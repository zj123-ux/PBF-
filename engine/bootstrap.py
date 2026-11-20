"""Excel 启动/恢复：确保 Excel 进程与 PBF 导航表处于可用状态。

一键执行和「运行中断后自动恢复」都走这里：
- 已就绪 → 直接返回
- 进程残留但 COM 断开 → 结束残留进程（kill_stale=True 时）
- 否则用真实 Excel 的 EXE 以文档参数启动（等价于双击），轮询到工作簿出现
"""
from __future__ import annotations

import logging
import subprocess
import threading
import time
from pathlib import Path

from .excel import ExcelEngine, ExcelError, reset_active_excel

log = logging.getLogger(__name__)

# 复用已打开的 Excel 时守候登录框的时长。AO 加载约 10s 后才弹登录框；
# 已用缓存凭据登录时框不会出现，等满 60s 会让每次复用都空等一分钟。
REUSE_LOGIN_WAIT_SEC = 15


def excel_pids() -> list[int]:
    """当前 Excel 进程 PID 列表（tasklist 解析）。"""
    try:
        out = subprocess.run(
            ["tasklist", "/FI", "IMAGENAME eq EXCEL.EXE", "/FO", "CSV", "/NH"],
            capture_output=True, text=True, timeout=15).stdout
    except Exception:
        return []
    pids: list[int] = []
    for line in out.splitlines():
        parts = [p.strip().strip('"') for p in line.split(",")]
        if len(parts) >= 2 and parts[0].upper() == "EXCEL.EXE" and parts[1].isdigit():
            pids.append(int(parts[1]))
    return pids


def kill_excel(pid: int | None = None) -> None:
    """结束 Excel 进程。

    pid 给定 → 只结束该进程（断线恢复的正确做法）；
    pid 为 None → 结束所有 EXCEL.EXE（仅保留给显式的人工命令，自动流程不要用）。
    """
    if pid is None:
        log.warning("结束所有 Excel 进程…")
        cmd = ["taskkill", "/IM", "EXCEL.EXE", "/F"]
    else:
        log.warning(f"结束 Excel 进程 PID={pid}…")
        cmd = ["taskkill", "/PID", str(pid), "/T", "/F"]
    try:
        subprocess.run(cmd, capture_output=True, text=True, timeout=30)
    except Exception as exc:
        log.warning(f"结束 Excel 失败（忽略）: {exc}")
    reset_active_excel()  # 进程已结束，清掉缓存的 COM 引用
    time.sleep(3)


class LoginHandle:
    """后台登录线程的结果句柄。

    outcome: None=线程仍在进行 | True=登录成功 | False=登录失败。
             判定登录是否完成只看 outcome（"登录成功信号"），不能拿
             "线程已启动/已结束"当成功——密码错误时线程照样退出，若被
             当成已登录，后续 COM 调用会被模态登录框永久阻塞。
    saw_box: 线程期间是否见过登录框。False+失败=从未弹框（可能已登录或
             AO 加载慢，应交同步登录再确认）；True+失败=密码错误或登录框
             卡住，应立即报错而不是继续跑。
    """

    def __init__(self) -> None:
        self.thread: threading.Thread | None = None
        self.outcome: bool | None = None
        self.saw_box = False
        self.error = ""

    def wait(self, timeout: float | None = None) -> bool | None:
        """等线程结束，返回 outcome（超时仍进行中返回 None）。"""
        if self.thread is not None:
            self.thread.join(timeout)
        return self.outcome

    @property
    def finished(self) -> bool:
        return self.outcome is not None


def _start_ao_login(flow_cfg: dict, expect_login: bool = True,
                    quiet: bool = False, timeout: float | None = None) -> LoginHandle | None:
    """后台线程自动登录 AO（返回结果句柄，未配置账号时返回 None）。

    AO 连接 BPC 时弹出的登录框是模态的，会阻塞 Excel 的 COM 消息泵，导致主线程的
    Workbooks 枚举卡住。这里用 uiautomation 在后台线程填登录框（不依赖 Excel COM），
    登录完成后模态框消失，主线程自然恢复。登录结果写入句柄的 outcome。

    quiet=True 用于「守候登录框」场景：等不到登录框属于正常情况，不告警。
    """
    try:
        from .orchestrator import login_config
        from .ao_login import AoLogin
    except Exception:
        return None
    login_cfg = login_config(flow_cfg)
    if not login_cfg.get("username"):
        return None
    wait = float(timeout if timeout is not None else login_cfg.get("timeout", 120))

    handle = LoginHandle()

    def _login() -> None:
        # uiautomation 需要在线程内初始化 COM（CoInitialize + 建 UIAutomation 对象），
        # 否则报 "尚未调用 CoInitialize"；用库自带的初始化器包一层。
        import uiautomation as auto
        with auto.UIAutomationInitializerInThread():
            try:
                ao = AoLogin(
                    username=login_cfg["username"],
                    password=login_cfg.get("password", ""),
                    client=str(login_cfg.get("client", "333")),
                    lang=str(login_cfg.get("lang", "zh")),
                )
                ok = ao.ensure(wait, expect_login=expect_login)
                handle.outcome = bool(ok)
                handle.saw_box = ao.saw_login_box
                if not ok and not quiet:
                    log.warning("后台自动登录未成功（超时或密码错误），请检查账号密码")
            except Exception as exc:
                handle.outcome = False
                handle.error = str(exc)
                if not quiet:
                    log.warning(f"后台自动登录异常: {exc}")

    t = threading.Thread(target=_login, daemon=True)
    handle.thread = t
    t.start()
    return handle


def _mark_login(engine: ExcelEngine, launched_now: bool, login_done: bool = False,
                login_handle: "LoginHandle | None" = None) -> ExcelEngine:
    """给引擎打上启动/登录状态，供 run_flow 判断是否还要再等登录框。

    launched_now:  本次是否由本函数启动了 Excel（AO 登录框可能还没弹出来）。
    login_done:    启动阶段已**确认**登录成功（或无需登录）——只有拿到登录
                   成功信号才能置 True；后台线程"已启动/已结束"不算数，
                   失败 outcome 由 run_flow 消费后给出明确报错。
    login_handle:  后台登录线程的结果句柄（未起线程时为 None）。
    """
    engine.launched_now = launched_now
    engine.login_done = login_done
    engine.login_handle = login_handle
    return engine


def ensure_excel_ready(flow_cfg: dict, kill_stale: bool = False,
                       stale_pid: int | None = None, wait_sec: int = 180):
    """确保真实 Excel 已运行且导航表已打开，否则自动打开。

    返回可复用的 ExcelEngine（所有成功路径），调用方可复用同一实例，避免重连到别的 Excel。
    打开导航表用真实 Excel 的 EXE 以文档参数启动（等价于双击文件），完整加载 AO 并连接
    BPC；若用 COM 的 Workbooks.Open，AO 常不初始化、进不了 PBF。
    """
    nav = flow_cfg.get("navigator") or {}
    pattern = nav.get("workbook_pattern", "PBF Pro - 全面预算系统流程导航*")
    path = Path(nav["launch_file"]) if nav.get("launch_file") else None

    # 1) 已经就绪（真实 Excel 里导航表已打开）——仅当确有 Excel 进程时才去连，
    #    否则 ExcelEngine 的 Dispatch 会拉起一个空 Excel，还得再关掉，反而制造竞态。
    if excel_pids():
        try:
            engine = ExcelEngine(timeout_sec=3)
            engine.find_workbook(pattern)
            # 复用已有 Excel：用户可能刚打开导航表、登录框还没弹出来。
            # 起一个后台线程守候登录框（没出现就按超时自行退出），
            # 这样随后卡在模态登录框上的导航表/刷新调用会被自动解锁。
            # 结果句柄交给 run_flow 等 outcome，不能把"线程已启动"当登录成功。
            # 超时只给 REUSE_LOGIN_WAIT_SEC：登录框 10s 左右就会弹，
            # 已登录时不会弹，等满 60s 纯属空等。
            handle = _start_ao_login(flow_cfg, quiet=True,
                                     timeout=REUSE_LOGIN_WAIT_SEC)
            return _mark_login(engine, launched_now=False, login_handle=handle)
        except ExcelError:
            pass

    # 2) 清理无窗口的残留实例（有 Excel 进程但没打开导航表；COM 可能连到它，
    #    后续宏会全部失败）。没有 Excel 进程时直接跳过，避免 Dispatch 又启动一个。
    if excel_pids():
        try:
            engine = ExcelEngine(timeout_sec=3)
            if engine.workbook_count() == 0:
                engine._excel.Quit()
                reset_active_excel()
                log.warning("已关闭无窗口的残留 Excel 实例")
                time.sleep(3)
        except ExcelError:
            pass

    # 3) 断开重连时：残留进程可能已经僵死，只结束本工具连接的那个 PID。
    #    定位不到目标时宁可跳过，也不要 taskkill 掉用户其它 Excel（可能有未保存内容）。
    if kill_stale:
        running = excel_pids()
        if stale_pid and stale_pid in running:
            kill_excel(stale_pid)
        elif running:
            log.warning(
                f"未定位到需要结束的 Excel 进程（目标 PID={stale_pid}），"
                f"跳过强制结束，避免影响其它 Excel 窗口")

    # 4) 打开导航表
    #    source=bw（默认）：启动空白 Excel → 文件 → Analysis → 从 BW 平台现取导航表。
    #      本地 .xlsb 会带缓存的连接/版本上下文，切版本容易出错；BW 每次都是新副本。
    #    source=file：直接打开 flow.yaml 里配置的本地文件（旧方式，保留兜底）。
    source = str(nav.get("source") or "bw").lower()
    if source == "bw":
        from .bw_open import open_navigator_from_bw

        engine = open_navigator_from_bw(flow_cfg, wait_sec=wait_sec)
        # bw_open 内部检测到登录框时会自动填写并等它关闭，走到这里说明登录已处理
        return _mark_login(engine, launched_now=True, login_done=True)

    if not path or not path.exists():
        raise ExcelError(
            f"导航表未打开，且路径无效：{path}；请检查 config/flow.yaml 的 navigator.launch_file")
    excel_exe = ExcelEngine._real_excel_path()
    if not excel_exe:
        raise ExcelError("找不到真实 Excel（EXCEL.EXE），无法打开导航表，请确认 Office 已安装")
    log.info(f"正在打开导航表：{path}")
    reset_active_excel()
    try:
        subprocess.Popen([excel_exe, str(path)])
    except Exception as exc:
        raise ExcelError(f"启动 Excel 打开导航表失败：{exc}") from exc
    # 后台登录：模态登录框会阻塞 Excel COM，用 uiautomation 线程先把它填掉
    login_handle = _start_ao_login(flow_cfg)
    # 等 Excel 进程起来（避免下面的 Dispatch 又拉起第二个实例），再连 COM
    time.sleep(3)
    engine = ExcelEngine(timeout_sec=20)
    try:
        engine._excel.DisplayAlerts = False
    except Exception:
        pass

    deadline = time.time() + wait_sec
    while time.time() < deadline:
        time.sleep(3)
        try:
            engine.find_workbook(pattern)
            log.info("导航表已打开，等待 AO 加载…")
            time.sleep(10)
            if login_handle is not None:
                # 等登录出结果；失败不在这里吞掉——run_flow 看句柄 outcome
                # 决定跳过登录还是明确报错（密码错误时"线程已结束"≠登录成功）
                login_handle.wait(timeout=90)
            return _mark_login(engine, launched_now=True, login_handle=login_handle)
        except ExcelError:
            continue
    raise ExcelError(f"等待导航表打开超时（{wait_sec}s）")
