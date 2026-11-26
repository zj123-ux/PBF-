"""PbfFlow — 按业务顺序编排 PBF 滚动预测全流程。

顺序：价格销量 → 客户费用 → HQ毛上(含2个运行包) → 其他成本 → 费用损益(含1个运行包) → 损益核对

所有界面元素位置均已现场验证（2026-11-12）：
- 导航表"打开"按钮固定在 AT52
- 运行包 Execute=Trigger_PS / Refresh=Refresh_Sheet，状态在 AO73
- 损益指标行固定在 57/69/86/172/187，期间列按表头动态解析
  （目标期间默认从数据目录名 data_YYYY.MM 推导，需要折叠的年度也自动推导）
"""
from __future__ import annotations

import fnmatch
import json
import logging
import os
import re
import time
from pathlib import Path
from typing import Any, Optional

import pandas as pd
import yaml

from .business_units import (NavRowError, nav_rows_from_grid, resolve_nav_row)
from .excel import ExcelEngine, ExcelError, is_disconnect
from .bootstrap import ensure_excel_ready, launch_no_box_ok_sec, wait_ao_login_idle
from .cancel import CancelledError, is_cancelled, raise_if_cancelled
from .header_periods import normalize_period_header
from .navigator import Navigator
from .writer import FormWriter
from paths import app_root

log = logging.getLogger(__name__)

CREDENTIALS_FILE = app_root() / "config" / "credentials.yaml"


def login_config(flow_cfg: dict) -> dict:
    """登录配置。优先级：环境变量 > config/credentials.yaml > flow.yaml 的 login 段。

    凭据不要写进 flow.yaml（会随项目一起传阅），放本机 credentials.yaml 或环境变量：
        PBF_USERNAME / PBF_PASSWORD / PBF_CLIENT / PBF_LANG
    """
    cfg = dict(flow_cfg.get("login") or {})
    if CREDENTIALS_FILE.exists():
        try:
            local = yaml.safe_load(CREDENTIALS_FILE.read_text(encoding="utf-8")) or {}
            cfg.update({k: v for k, v in local.items() if v not in (None, "")})
        except Exception as exc:
            log.warning(f"读取 {CREDENTIALS_FILE} 失败: {exc}")
    for key, env in (("username", "PBF_USERNAME"), ("password", "PBF_PASSWORD"),
                     ("client", "PBF_CLIENT"), ("lang", "PBF_LANG")):
        value = os.environ.get(env)
        if value:
            cfg[key] = value
    return cfg


def _period_index(period: str) -> Optional[int]:
    """'2027.03' → 月序号（便于比较先后）。"""
    m = re.fullmatch(r"(\d{4})\.(\d{2})", (period or "").strip())
    return int(m.group(1)) * 12 + int(m.group(2)) - 1 if m else None


def _num(v) -> Optional[float]:
    """尽量转数值；空串/文本/None → None。"""
    try:
        return float(v)
    except (TypeError, ValueError):
        return None


def _blank(v) -> bool:
    return v is None or (isinstance(v, str) and not v.strip())


def _norm_version(text) -> str:
    """版本名归一化：口语「财总会版」与 AO「财总版」视为同一个。"""
    return re.sub(r"\s+", "", str(text or "")).replace("财总会版", "财总版")


def _version_matches(expected, actual) -> bool:
    e, a = _norm_version(expected), _norm_version(actual)
    return bool(e and a and (e in a or a in e))


def _lock_user(message: str) -> str:
    """从「表单被XXX用户锁定」里提取用户名；提取不到返回空串。"""
    text = str(message or "")
    # 用户名可能是中文、英文 ID，也可能带空格（Zhang San）
    m = re.search(r"被\s*(.{1,40}?)\s*(?:用户)?\s*锁定", text)
    if m:
        return m.group(1).strip()
    m = re.search(r"locked\s+by\s+([^,.;\s]{1,40})", text, re.I)
    return m.group(1).strip() if m else ""


class FormLockedError(ExcelError):
    """表单被其他用户锁定：跳过该表单并结束流程（不是普通失败，不重试）。"""

    def __init__(self, form_id: str, label: str, message: str) -> None:
        self.form_id = form_id
        self.label = label
        self.raw_message = " ".join(str(message or "").split())
        self.user = _lock_user(self.raw_message)
        if self.user:
            text = (f"{label}表单被{self.user}用户锁定，已跳过{label}编辑并结束流程。"
                    f"原始提示：{self.raw_message}")
        else:
            text = (f"{label}表单疑似被其他用户锁定（{self.raw_message}），"
                    f"已跳过{label}编辑并结束流程。")
        super().__init__(text)


class SessionExpiredError(ExcelError):
    """AO/BPC 会话失效。

    与「Excel 断开」不同：进程和 COM 都还在，但导航表宏统一报 0x800A9C68、
    右侧步骤面板不再随选中刷新、版本被重置回月初版。只有重开 Excel 重新
    登录才能恢复，所以 run_with_recovery 会强制重启而不是复用原实例。
    """


# AO 宏（Execute/Refresh_Status）在会话过期时统一抛的特征码
_SESSION_EXPIRED_MARKERS = ("0x800A9C68", "-2146788248")


def is_session_expired(exc: BaseException) -> bool:
    """错误文本里是否带 AO/BPC 会话失效的特征码。

    只认 0x800A9C68 / -2146788248，不认「发生意外」/ -2147352567——那俩是
    通用 DISP_E_EXCEPTION 包装，版本与期间不匹配等普通错误也长这样，
    拿它们触发强制重启会误杀正常会话。
    """
    text = str(exc or "")
    return any(marker in text for marker in _SESSION_EXPIRED_MARKERS)


def _package_items(packages) -> list[tuple[int, str]]:
    """steps[].after_close.packages → [(行号, 期望包名)]。

    兼容两种写法：
    - ``[52, 53]``      旧行号列表：无包名校验
    - ``52: 收入计算``  行号 → 期望包名映射：开单后核对清单 AB73 步骤名，
      不符即中止（步骤面板停在别的同类型行上时会开到错误的包上下文）

    映射写法里包名必须非空：空的会悄悄跳过 AB73 校验，等于误配把新
    防线关掉了——宁可在跑之前就报错。
    """
    if isinstance(packages, dict):
        items: list[tuple[int, str]] = []
        for r, n in packages.items():
            name = str(n or "").strip()
            if not name:
                raise ExcelError(
                    f"after_close.packages 配置了「行号: 包名」映射，但行 {r} 的包名为空"
                    "（空包名会跳过 AB73 步骤名校验）。请补上包名，"
                    "或改回行号列表写法 [52, 53]")
            items.append((int(r), name))
        return items
    return [(int(r), "") for r in (packages or [])]


class PbfFlow:
    def __init__(self, engine: ExcelEngine, flow_cfg: dict, form_cfgs: dict, data_dir: Path,
                 persist_panel_cache: bool = False,
                 user_key: Optional[str] = None) -> None:
        self.engine = engine
        self.cfg = flow_cfg
        self.nav_cfg = flow_cfg["navigator"]
        self.pkg_cfg = flow_cfg["package"]
        self.wb_cfg = flow_cfg.get("verify_writeback") or {}
        self.form_cfgs = form_cfgs
        self.data_dir = Path(data_dir)
        # 导航表配置（含 dimensions.yaml 的版本单元格）；run_flow 会补充 dims_nav
        self.nav_dims_cfg = dict(self.nav_cfg)
        # 版本|nav_cell → 上次验证过的步骤面板指纹。
        # 用于区分"面板已是该行的正确内容"（不必刷新）和"面板没刷新"（会话失效）。
        # 跨进程持久化：上一次运行可能刚好把面板停在目标行，新进程若无法识别，
        # 会把"内容没变化"误判成会话失效，触发不必要的 Excel 重启和重新登录。
        self._panel_cache: dict[str, str] = {}
        self.persist_panel_cache = persist_panel_cache
        self._panel_cache_file = app_root() / "logs" / "nav_panel_cache.json"
        if persist_panel_cache:
            self._load_panel_cache()
        # 面板指纹缓存按账号隔离：换账号登录时导航表按权限重新生成，行号可能
        # 整体移动；上一个账号验证过的指纹不能当作当前账号的正确状态。
        # user_key 仅用于测试注入；正常运行取配置/env 里的登录用户名。
        self.login_user = (str(user_key).strip() if user_key is not None
                           else str(login_config(flow_cfg).get("username") or "").strip())
        # 导航标签日志去重（同一格每轮只提示一次）
        self._nav_label_logged: set[str] = set()
        self._nav_label_warned: set[str] = set()
        self._nav_ambiguous_warned: set[str] = set()
        # 每张表单打开后要核对的版本（由 run_flow 根据 --version-key/--version 设置）
        self.form_version_key: Optional[str] = None
        self.form_version_display: Optional[str] = None
        # read_indicators 是否读到目标期间指标全 0（清 0 模式放行时置 True）
        self.indicators_all_zero = False

        # 目标期间：优先 verify.target_period，否则从数据目录名推导。
        # 目录名可能含多个期间（如 data_2027.01_2027.04），取最大期间，保证损益核对
        # 与版本窗口校验覆盖到最后一个导入期间。
        explicit = str((flow_cfg.get("verify") or {}).get("target_period") or "").strip()
        targets = [f"{y}.{m}" for y, m in re.findall(r"(\d{4})\.(\d{2})", self.data_dir.name)]
        self.target_period = explicit or (max(targets, key=_period_index) if targets else "")
        if self.target_period:
            log.info(f"目标期间: {self.target_period}")
        else:
            log.warning(f"未能从数据目录 '{self.data_dir.name}' 推导目标期间，将展示全部期间列")

    # ---------- 基础 ----------

    def _nav_wb(self):
        return self.engine.get_workbook(self.nav_cfg["workbook_pattern"], wait_sec=60)

    def _nav_ws(self):
        return self.engine.open_sheet(self._nav_wb(), self.nav_cfg["sheet"])

    def _read(self, ws, cell: str) -> Any:
        # AO 计算期间会短暂拒绝调用（RPC_E_SERVERCALL_RETRYLATER），统一重试
        return self.engine._retry(lambda: self.engine.read(ws, cell))

    def _load_data(self, path: Path) -> list[dict]:
        if not path.exists():
            raise ExcelError(f"数据文件不存在: {path}")
        if path.suffix.lower() == ".csv":
            df = pd.read_csv(path)
        else:
            df = pd.read_excel(path)
        return df.to_dict(orient="records")

    def _close_if_open(self, pattern: str) -> None:
        try:
            wb = self.engine._retry(lambda: self.engine.find_workbook(pattern), attempts=4, delay=1.5)
        except ExcelError:
            return
        name = wb.Name                      # 关闭后对象即断开，先取名字
        try:
            self.engine._retry(lambda: self.engine.close_workbook(wb, save=False), attempts=3, delay=1.5)
        except Exception as exc:
            log.warning(f"关闭 {name} 失败（忽略）: {exc}")
            return
        log.info(f"已关闭旧实例: {name}")
        time.sleep(2)

    # ---------- 导航操作 ----------

    def confirm_version(self, expected: Optional[str] = None) -> str:
        ws = self._nav_ws()
        v = self._read(ws, self.nav_cfg["version_cell"])
        v = "" if v is None else str(v)
        if expected and not _version_matches(expected, v):
            raise ExcelError(f"版本不符：当前 '{v}'，期望包含 '{expected}'")
        log.info(f"版本确认: {v}")
        return v

    def ensure_navigator_version(self) -> str:
        """导航表版本必须与表单目标版本一致。

        长流程中 AO 会话掉线后，导航表会被重置回默认的月初版；此时继续
        打开运行包/损益表会拿到错误版本的数据。发现漂移就切回并刷新。
        """
        expected = self.form_version_display
        key = self.form_version_key
        if not expected and not key:
            return ""
        ws = self._nav_ws()
        current = str(self._read(ws, self.nav_cfg["version_cell"]) or "")
        if expected and _version_matches(expected, current):
            return current

        log.warning(f"导航表版本漂移：当前 '{current}'，目标 '{expected or key}'，重新切换")
        if not key:
            raise SessionExpiredError(
                f"导航表版本为 '{current}'，期望 '{expected}'，但没有版本成员键，"
                "无法自动切回；请重新登录 AO 后重试")
        try:
            from .navigator import Navigator

            Navigator(self.engine, self.nav_dims_cfg).switch_version(key, expected)
            self.refresh_flow_retry(attempts=2, delay=5)
        except SessionExpiredError:
            raise
        except ExcelError as exc:
            raise SessionExpiredError(f"导航表版本切回失败：{exc}") from exc

        current = str(self._read(ws, self.nav_cfg["version_cell"]) or "")
        if expected and not _version_matches(expected, current):
            raise SessionExpiredError(
                f"导航表版本切回后仍为 '{current}'（目标 '{expected}'），"
                "AO/BPC 会话可能已失效")
        log.info(f"导航表版本已切回: {current}")
        return current

    def ensure_form_version(self, form_id: str) -> None:
        """打开表单后核对其预算版本；不是目标版本就写入成员键并点「刷新表单」。

        背景：从月初版快捷方式进入 PBF 时，新打开的表单可能仍是月初版。
        表单版本区：F4=显示文本，E4=目标成员键，G4=SAP 当前变量值，H4=是否一致。
        实测：写 E4 后 F4 立即变化，执行 Refresh_Sheet 后 G4/H4 才更新。
        """
        cfg = self.form_cfgs[form_id]
        disp_cell = cfg.get("version_cell")
        key = self.form_version_key
        expected = self.form_version_display
        if not disp_cell or (not key and not expected):
            return

        pattern = cfg["workbook_pattern"]
        key_cell = cfg.get("version_key_cell", "E4")
        cur_cell = cfg.get("version_current_cell", "G4")
        match_cell = cfg.get("version_match_cell", "H4")

        def state():
            wb = self.engine.get_workbook(pattern, wait_sec=30)
            ws = self.engine.open_sheet(wb, cfg["sheet"])
            vals = {
                "display": self._read(ws, disp_cell),
                "key": self._read(ws, key_cell) if key_cell else None,
                "current": self._read(ws, cur_cell) if cur_cell else None,
                "match": self._read(ws, match_cell) if match_cell else None,
            }
            return wb, ws, vals

        def is_ok(vals: dict) -> bool:
            if expected and not _version_matches(expected, vals["display"]):
                return False
            if key:
                if key_cell and str(vals["key"] or "").strip() != key:
                    return False
                if cur_cell and str(vals["current"] or "").strip() != key:
                    return False
            if match_cell and str(vals["match"] or "").strip() not in ("", "Y"):
                return False
            return True

        wb, ws, vals = state()
        if is_ok(vals):
            log.info(f"{form_id} 版本正确: {vals['display']}")
            return
        if not key:
            raise ExcelError(
                f"{form_id} 版本不符：当前 '{vals['display']}'，"
                f"期望 '{expected}'，但未提供版本成员键，无法自动切换")
        if not key_cell:
            raise ExcelError(f"{form_id} 未配置 version_key_cell，无法自动切换版本")

        log.warning(
            f"{form_id} 版本为 '{vals['display']}'，切换到 '{expected or key}' ({key}) 并刷新表单")
        self.engine.write(ws, key_cell, key)

        macro = (cfg.get("macros") or {}).get("refresh") or "Refresh_Sheet"
        try:
            self.engine.run_macro(macro, wb)
        except ExcelError as exc:
            raise ExcelError(f"{form_id} 刷新表单失败，无法切换版本: {exc}") from exc

        deadline = time.time() + 90
        last = vals
        while time.time() < deadline:
            time.sleep(2)
            try:
                _wb, _ws, last = state()
            except ExcelError as exc:
                log.warning(f"{form_id} 刷新后暂时读不到表单，继续等待: {exc}")
                continue
            if is_ok(last):
                log.info(f"{form_id} 版本已切换并刷新: {last['display']}")
                return
        raise ExcelError(
            f"{form_id} 切版本未生效：显示 '{last['display']}'，"
            f"目标键 {key}，当前值 '{last['current']}'，一致性 '{last['match']}'。"
            "请在 AO 中手动选择版本后重试")

    def refresh_flow(self) -> None:
        wb = self._nav_wb()
        try:
            self.engine.run_macro(self.nav_cfg["refresh_flow_macro"], wb)
        except ExcelError as exc:
            text = str(exc)
            # "发生意外"（0x80020009 / 内层 0x800A9C68）几乎都来自 AO/BPC 服务端：
            # 会话过期，或版本与目标期间不匹配。给一句能直接下手的人话。
            if ("发生意外" in text or "-2147352567" in text
                    or is_session_expired(exc)):
                raise SessionExpiredError(
                    "刷新流程（Refresh_Status）失败：AO/BPC 会话可能已过期，或版本与目标期间不匹配。"
                    "请手动打开 Excel 的 Analysis 选项卡确认已连接 BPC 并重新登录；"
                    "若期间超出当前版本窗口，请先在 AO 里把版本滚动到覆盖该期间的版本后重试。"
                    f" 原始错误：{text}"
                ) from exc
            raise
        log.info("已刷新流程")

    def refresh_flow_retry(self, attempts: int = 3, delay: float = 5.0) -> None:
        """刷新流程，带重试。

        AO 刚登录/刚打开导航表时，Refresh_Status 常因服务端会话尚未就绪而失败；
        这是可恢复的瞬时状态，不该直接把整轮流程判死。
        """
        last: Optional[BaseException] = None
        for i in range(1, attempts + 1):
            try:
                self.refresh_flow()
                return
            except ExcelError as exc:
                last = exc
                if i < attempts:
                    log.warning(f"刷新流程失败（第 {i}/{attempts} 次），{delay:.0f}s 后重试：{exc}")
                    time.sleep(delay)
        raise last  # type: ignore[misc]

    def recover(self, session_expired: bool = False) -> bool:
        """Excel 断开/AO 会话失效后自动恢复：重开导航表 → 重新登录 → 换上新的引擎。

        session_expired=True 时必须强制重启 Excel：会话失效时进程还在、
        导航表也还开着，复用旧实例无法重建 BPC 会话。

        返回 True 表示已恢复，调用方可以重做失败的步骤；False 表示恢复失败。
        """
        rcfg = self.cfg.get("recovery") or {}
        if session_expired:
            log.warning("AO/BPC 会话已失效，重新打开 Excel 并登录以恢复…")
        else:
            log.warning("Excel 连接已断开，开始自动恢复…")
        # 只结束本流程连接的那个 Excel 进程；不碰用户同时打开的其它 Excel。
        stale_pid = getattr(self.engine, "pid", None)
        try:
            self.engine = ensure_excel_ready(
                self.cfg, kill_stale=bool(rcfg.get("kill_stale_excel", True)),
                stale_pid=stale_pid, force_restart=session_expired)
        except Exception as exc:
            log.error(f"自动恢复失败：{exc}")
            return False

        login_cfg = login_config(self.cfg)
        # 启动阶段的后台登录线程以结果句柄为准：成功才跳过登录，
        # 失败/未确认都要在这里重新同步登录，不能把"线程已结束"当已登录
        login_confirmed = bool(getattr(self.engine, "login_done", False)) or \
            await_background_login(self.engine, self.cfg,
                                   float(login_cfg.get("timeout", 120)) + 30)
        if login_cfg.get("username") and not login_confirmed:
            from .ao_login import AoLogin

            launched = bool(getattr(self.engine, "launched_now", False))
            login = AoLogin(
                username=login_cfg["username"],
                password=login_cfg.get("password", ""),
                client=str(login_cfg.get("client", "333")),
                lang=str(login_cfg.get("lang", "zh")),
            )
            ok = login.ensure(float(login_cfg.get("timeout", 120)),
                              expect_login=launched,
                              no_box_ok_after=launch_no_box_ok_sec(self.cfg) if launched else None)
            if not ok:
                log.error("恢复后自动登录失败（已停止自动重试，避免账号锁定），"
                          "请在 Excel 中手动登录后重试")
                return False
            if getattr(login, "saw_login_box", False) \
                    and not getattr(self.engine, "login_settled", False):
                # 同步登录真的填过登录框：加载项随后会自动刷新数据源，
                # 等它收尾再继续（同 wait_ao_login_idle 的屏障）
                wait_ao_login_idle(self.engine, self.cfg)
        elif login_confirmed:
            log.info("启动阶段已处理 AO 登录，恢复流程跳过重复登录")
        # AO 刚登录完导航表还在加载，等它可读再交还给调用方
        try:
            self.engine._retry(self._nav_wb, attempts=10, delay=2)
            time.sleep(3)
        except Exception as exc:
            log.warning(f"恢复后导航表暂不可读（继续尝试）: {exc}")
        # 重新登录后导航表会回到"全流程视图"（80+ 行），必须刷新状态才能恢复成配置里的 19 行
        try:
            self.refresh_flow_retry(attempts=3, delay=5)
        except Exception as exc:
            log.warning(f"恢复后刷新流程失败（继续尝试）: {exc}")
        log.info("自动恢复完成，可以继续执行")
        return True

    def check_version_window(self, version_text: str) -> None:
        """校验目标期间是否落在版本滚动窗口内。

        版本名形如「滚动预测2026年11月月初版」：起始期间 2026.11，窗口默认 6 个月。
        目标期间早于起始期间 → 直接报错；超出窗口 → 只告警（版本可能已滚动但名字没变）。
        """
        if not self.target_period:
            return
        m = re.search(r"(\d{4})\s*年\s*(\d{1,2})\s*月", version_text or "")
        if not m:
            log.warning(f"版本名解析不出起始期间：'{version_text}'，跳过期间窗口校验")
            return
        start = f"{m.group(1)}.{int(m.group(2)):02d}"
        tgt, base = _period_index(self.target_period), _period_index(start)
        if tgt is None or base is None:
            return
        if tgt < base:
            raise ExcelError(
                f"目标期间 {self.target_period} 早于版本起始期间 {start}"
                f"（当前版本：{version_text}），请先切换版本")
        window = int((self.cfg.get("verify") or {}).get("version_window_months") or 6)
        if window > 0 and tgt - base > window - 1:
            log.warning(
                f"目标期间 {self.target_period} 超出版本 {start} 起 {window} 个月的滚动窗口，"
                f"请确认版本是否已滚动")
        else:
            log.info(f"期间校验通过：版本 {start} 起，目标 {self.target_period}")

    # ---------- 右侧「步骤事项」面板 ----------

    def _panel_range(self) -> str:
        return str(self.nav_cfg.get("panel_range") or "AP48:AU60")

    def _panel_fingerprint(self) -> str:
        """面板指纹。面板是 VBA 按选中的业务单元静态写入的，选中后必须变化。

        会话掉线时 VBA 宏会静默失败，面板停在上一张表的清单——此时 AT52
        的文案可能仍是"打开"，点下去却打开错误的东西或毫无反应。
        """
        ws = self._nav_ws()
        rng = self._panel_range()
        try:
            vals = self.engine._retry(lambda: ws.Range(rng).Value2)
        except Exception as exc:
            raise ExcelError(f"读取导航步骤面板 {rng} 失败: {exc}") from exc
        if not isinstance(vals, tuple):
            vals = ((vals,),)
        elif vals and not isinstance(vals[0], tuple):
            vals = (vals,)
        flat: list[str] = []
        for row in vals:
            for v in row:
                flat.append("" if v is None else str(v).strip())
        return "|".join(flat)

    def _panel_cache_key(self, nav_cell: str) -> str:
        """面板指纹按「账号|版本|导航格」区分。

        - 账号不同：导航表按权限重新生成，行号可能整体移动，缓存不能跨账号复用；
        - 版本不同：步骤清单可能不同。
        """
        ver = self.form_version_display or self.form_version_key or ""
        user = self.login_user or "-"
        return f"{user}|{ver}|{nav_cell}"

    def _load_panel_cache(self) -> None:
        try:
            data = json.loads(self._panel_cache_file.read_text(encoding="utf-8"))
        except FileNotFoundError:
            return
        except Exception as exc:
            log.debug(f"读取面板缓存失败（忽略）: {exc}")
            return
        if isinstance(data, dict):
            self._panel_cache.update(
                {str(k): v for k, v in data.items() if isinstance(v, str)})

    def _save_panel_cache(self) -> None:
        if not self.persist_panel_cache:
            return
        try:
            self._panel_cache_file.parent.mkdir(parents=True, exist_ok=True)
            self._panel_cache_file.write_text(
                json.dumps(self._panel_cache, ensure_ascii=False, indent=1),
                encoding="utf-8")
        except Exception as exc:
            log.debug(f"保存面板缓存失败（忽略）: {exc}")

    def _cell_text(self, ws, addr: str) -> str:
        """读取单元格文本；合并单元格取左上角，读不到返回空串。"""
        if not addr:
            return ""

        def _read():
            rng = ws.Range(addr)
            try:
                v = rng.MergeArea.Cells(1, 1).Value2
            except Exception:
                v = rng.Value2
            if isinstance(v, tuple) and len(v) == 1 and isinstance(v[0], tuple):
                v = v[0][0]
            return str(v or "").strip()
        try:
            return self.engine._retry(_read, attempts=5, delay=1)
        except Exception as exc:
            log.warning(f"读取导航标签 {addr} 失败: {exc}")
            return ""

    def _nav_label(self, ws, nav_cell: str) -> str:
        """导航行身份标签：AA(业务模块) + AB(流程步骤) + AC(业务单元)。

        人工判断表单位置就是看这三列：AC 相同的行（如 AC65/AC67 都是
        「CBG-智屏-乐华品牌」）必须靠 AB 的「编制/查询」区分，所以三者缺一不可。
        """
        row = "".join(ch for ch in nav_cell if ch.isdigit())
        if not row:
            return self._cell_text(ws, nav_cell)
        parts = [
            self._cell_text(ws, f"AA{row}"),
            self._cell_text(ws, f"AB{row}"),
            self._cell_text(ws, nav_cell),
        ]
        return " | ".join(p for p in dict.fromkeys(parts) if p)

    def _expect_labels(self, expect_labels: Optional[list[str]]) -> list[str]:
        """归一化标签关键字；None = 用单元级 nav_expect。"""
        if expect_labels is None:
            expect_labels = (self.cfg.get("_unit") or {}).get("nav_expect") or []
        return [str(x).strip() for x in expect_labels if str(x).strip()]

    def _nav_scan_range(self) -> str:
        return str(self.nav_cfg.get("locate_scan_range") or "AA1:AC200")

    def _nav_grid(self, ws) -> list[dict]:
        """批量只读扫描导航表 AA:AC，返回可定位的行。

        读失败时返回空列表：调用方退回固定行，单格标签校验仍是最后一道防线。
        """
        rng = self._nav_scan_range()
        try:
            vals = self.engine._retry(lambda: ws.Range(rng).Value2,
                                      attempts=5, delay=1)
        except Exception as exc:
            log.warning(f"读取导航表 {rng} 失败，改用固定行定位：{exc}")
            return []
        start = int("".join(ch for ch in rng.split(":")[0] if ch.isdigit()) or 1)
        return nav_rows_from_grid(vals, start)

    def _resolve_nav_cell(self, ws, nav_cell: str,
                          expect_labels: Optional[list[str]]) -> str:
        """按 AA+AB+AC 标签内容定位实际导航格（换账号行号移动也能对上）。

        - 唯一命中 → 用命中行；
        - 多行命中且配置行在其中 → 沿用配置行并告警（与只有单格校验时一致）；
        - 零命中 / 多行命中但配置行不在其中 → 中止，不选中、不打开任何表单。
        """
        if not bool(self.nav_cfg.get("locate_by_label", True)):
            return nav_cell
        expect = self._expect_labels(expect_labels)
        if not expect:
            return nav_cell
        rows = self._nav_grid(ws)
        if not rows:
            return nav_cell
        try:
            cell, mode = resolve_nav_row(rows, nav_cell, expect)
        except NavRowError as exc:
            label = self._nav_label(ws, nav_cell)
            raise ExcelError(
                f"{nav_cell} 按标签定位失败：{exc}。"
                f"配置行当前标签是「{label or '空'}」。"
                "当前账号的导航表布局可能与配置不一致，已中止，"
                "未选中任何行、未打开任何表单。请用该账号核对 "
                "config/business_units.yaml 的 nav_cell / nav_expect"
                "（可先运行 python cli.py nav-scan --unit <单元> 查看）")
        if mode == "unique" and cell != nav_cell:
            log.warning(
                f"导航行按标签定位：配置 {nav_cell} → 实际 {cell}"
                f"（标签「{self._nav_label(ws, cell)}」）。"
                "请核对配置：换账号后行号可能整体移动，"
                "也可能是 nav_expect 指向了另一行")
        elif mode == "ambiguous" and nav_cell not in self._nav_ambiguous_warned:
            self._nav_ambiguous_warned.add(nav_cell)
            log.warning(
                f"{nav_cell} 的标签关键字 {expect} 匹配多行，沿用配置行；"
                "建议补 AB 关键字（如「编制」「查询」）以便按内容唯一定位")
        return cell

    def _check_nav_label(self, ws, nav_cell: str,
                         expect_labels: Optional[list[str]]) -> None:
        """选中前核对导航行标签，防止账号权限不同导致固定行指向别的单元。

        只做读取，不选中、不触发 SelectionChange——标签不对时零副作用中止。
        未配置关键字时只告警一次（兼容旧配置，但等于放弃了这层保护）。
        """
        expect = [str(x).strip() for x in (expect_labels or []) if str(x).strip()]
        if not expect:
            if nav_cell not in self._nav_label_warned:
                self._nav_label_warned.add(nav_cell)
                log.warning(
                    f"未配置 nav_expect，跳过 {nav_cell} 的导航标签校验——"
                    "换账号后导航行错位将无法发现")
            return
        label = self._nav_label(ws, nav_cell)
        missing = [k for k in expect if k not in label]
        if missing:
            raise ExcelError(
                f"导航行 {nav_cell} 的标签（AA+AB+AC）是「{label or '空'}」，"
                f"未包含 {missing}。"
                "当前账号的导航表布局可能与配置不一致（不同权限的行会整体移动），"
                "已中止，未打开任何表单。请用该账号核对 config/business_units.yaml "
                "的 nav_cell / nav_expect")
        if nav_cell not in self._nav_label_logged:
            self._nav_label_logged.add(nav_cell)
            log.info(f"{nav_cell} 标签校验通过：{label}")

    def _select_nav_cell(self, ws, nav_cell: str,
                         expect_labels: Optional[list[str]] = None) -> None:
        """选中业务单元行（先核对标签）。

        SelectionChange 只在选中格**变化**时触发；重试时若仍选着同一格，
        面板不会再刷新（这也是 AT52 点下去没反应的原因之一）。先点 A1
        强制一次变化，再选目标格。
        """
        if expect_labels is None:
            expect_labels = (self.cfg.get("_unit") or {}).get("nav_expect")
        self._check_nav_label(ws, nav_cell, expect_labels)
        try:
            ws.Range("A1").Select()
        except Exception:
            pass
        self.engine._retry(lambda: ws.Range(nav_cell).Select(), attempts=5, delay=2)

    def _row_type(self, ws, btn: str) -> str:
        """按钮所在行对应的「行类型」列（AR 列）值：表单 / 运行包。"""
        row = "".join(ch for ch in btn if ch.isdigit())
        return str(self._read(ws, f"AR{row}") or "").strip() if row else ""

    def _static_panel_acceptable(self, ws, nav_cell: str, want: Optional[str],
                                 row_type: str) -> bool:
        """面板内容没变化时，判断它是否本来就该停在这一行。

        不同导航行可以共用同一套流程步骤模板——现场实测「CBG-商用」与
        「CBG-乐华」的 CBG损益查询流程返回同一个步骤模板，选中后步骤面板
        内容一字不差。此时"面板必须变化"的判定会误报会话失效（复核时
        Refresh_Status 仍能正常刷新，会话其实活着）。

        导航表自己的 Worksheet_SelectionChange 会把本次选中行写进 B4
        （「本次选择行数」），且写在任何 AO 调用之前——读它就能确认这次
        选中确实被导航表处理过，再按行类型接受静止面板。B4 不可读时退回
        原来的行类型放行，兼容没有该单元格的导航表。
        """
        if want:
            if not row_type or row_type != want:
                return False
        elif row_type not in ("表单", "运行包"):
            return False
        try:
            raw = self._read(ws, "B4")
        except Exception:
            raw = None
        if raw in (None, ""):
            log.info(f"{nav_cell} 面板内容未变化，导航表 B4 不可读，按行类型「{row_type}」接受")
            return True
        try:
            selected = int(float(raw))
        except (TypeError, ValueError):
            log.info(f"{nav_cell} 面板内容未变化，B4={raw!r} 不是行号，按行类型「{row_type}」接受")
            return True
        target = int("".join(ch for ch in nav_cell if ch.isdigit()) or 0)
        if target and selected == target:
            log.info(f"{nav_cell} 面板内容与上一行相同（共用流程模板），"
                     f"B4={selected} 确认本次选中已被导航表处理，按行类型「{row_type}」接受")
            return True
        log.warning(f"{nav_cell} 面板内容未变化，但导航表 B4={selected} 不是本次选中行"
                    f"（{target}）——选中可能未生效或会话异常，不能按静止面板接受")
        return False

    def _wait_panel_refresh(self, ws, nav_cell: str, btn: str,
                            before: str, timeout_sec: float,
                            want: Optional[str] = None,
                            accept_static: bool = False) -> tuple[bool, str, bool]:
        """等面板刷新出"打开"按钮；返回 (是否刷新, 按钮当前文案, 是否命中缓存)。

        命中缓存（面板内容 == 该行上次验证过的正确指纹）也算刷新成功：
        上一次运行可能刚好把面板停在这一行，选中前后内容一致属正常。

        accept_static：无缓存时（新进程冷启动），刷新流程已成功过一轮之后，
        面板内容没变也可接受——要求按钮行行类型匹配，且导航表 B4 记录的
        本次选中行就是目标行（见 _static_panel_acceptable）。否则会把
        "面板本来就停在目标行"或"两行共用同一流程模板"误判成会话失效。
        """
        deadline = time.time() + max(float(timeout_sec), 1.0)
        cached = self._panel_cache.get(self._panel_cache_key(nav_cell))
        btn_text = ""
        while time.time() < deadline:
            time.sleep(1)
            try:
                now = self._panel_fingerprint()
                btn_text = str(self._read(ws, btn) or "").strip()
                row_type = self._row_type(ws, btn)
            except ExcelError as exc:
                log.warning(f"等待 {nav_cell} 面板时读取失败（继续等）: {exc}")
                continue
            if "打开" in btn_text:
                if now != before:
                    return True, btn_text, False
                if cached is not None and now == cached:
                    return True, btn_text, True
                if accept_static and self._static_panel_acceptable(
                        ws, nav_cell, want, row_type):
                    return True, btn_text, True
        return False, btn_text, False

    def open_via_nav(self, nav_cell: str, open_cell: Optional[str] = None,
                     expect: str = "form", wait_sec: Optional[int] = None,
                     pattern: Optional[str] = None,
                     expect_labels: Optional[list[str]] = None) -> str:
        """选中业务单元行 → 等面板刷新 → 点"打开" → 返回新打开的工作簿名。

        expect: 'form' 期望打开编制表；'package' 期望打开运行包清单。
        pattern: 编制表的 workbook_pattern（有些表单名里没有"编制"二字，只能按模式匹配）。
        expect_labels: 导航行标签必须包含的关键字；None = 用单元级 nav_expect。
        """
        wb = self._nav_wb()
        ws = self.engine.open_sheet(wb, self.nav_cfg["sheet"])
        btn = open_cell or self.nav_cfg["open_button_cell"]
        panel_wait = float(self.nav_cfg.get("panel_wait_sec") or 15)

        # 先按标签内容定位实际行（只读），再进行选中/面板刷新。
        # 定位失败会在选中前中止，保证零副作用。
        expect_labels = self._expect_labels(expect_labels)
        nav_cell = self._resolve_nav_cell(ws, nav_cell, expect_labels)

        # "打开"靠导航表的 SelectionChange 触发，偶尔会被 Excel/AO 忙吞掉，因此重试几轮
        want = {"form": "表单", "package": "运行包"}.get(expect)
        last_names: list[str] = []
        via_cache = False
        for attempt in range(1, 4):
            self.engine._retry(lambda: (wb.Activate(), ws.Activate()), attempts=5, delay=2)
            before = set(self.engine._retry(
                lambda: self.engine.workbook_names(), attempts=5, delay=2))
            before_panel = self._panel_fingerprint()
            self._select_nav_cell(ws, nav_cell, expect_labels)
            # 第 1 轮不允许"面板没动就算刷新过"——那可能真是会话死了，
            # 交给下面的刷新流程去分辨；刷新成功过一轮后才接受静止面板。
            fresh, btn_text, via_cache = self._wait_panel_refresh(
                ws, nav_cell, btn, before_panel, panel_wait,
                want=want, accept_static=attempt > 1)

            if not fresh:
                # 内容没变不一定是会话失效：商用/乐华等共用同一流程模板的行，
                # 选中后步骤面板本来就一字不差（B4 会确认选中已生效）。先刷新流程，
                # 刷新成功后再按 B4/行类型判断，避免每轮第一次核对都误报过期。
                log.warning(
                    f"选中 {nav_cell} 后步骤面板内容未变化（{btn}='{btn_text}'），"
                    "可能该行与上一行共用同一流程模板；先刷新流程确认会话状态")
                if attempt == 1:
                    # 先走工具自带的「刷新流程」：它正好能区分瞬时忙和会话过期
                    try:
                        self.refresh_flow_retry(attempts=2, delay=5)
                    except SessionExpiredError:
                        raise
                    except ExcelError as exc:
                        raise SessionExpiredError(
                            f"选中 {nav_cell} 后面板未刷新，刷新流程也失败：{exc}") from exc
                    continue
                raise SessionExpiredError(
                    f"选中 {nav_cell} 后步骤面板一直没有刷新（{btn}='{btn_text}'）。"
                    "AO/BPC 会话可能已过期，需要重新打开 Excel 并登录；"
                    "请确认 AO 已连接 BPC 后重试。")

            if "打开" not in btn_text:
                raise ExcelError(
                    f"{nav_cell} 的步骤面板已刷新，但 {btn} 不是「打开」（实际 '{btn_text}'）。"
                    f"请检查该步骤的行号配置（当前 {btn}）与面板内容。")

            row_type = self._row_type(ws, btn)
            if want and row_type and row_type != want:
                names = []
                for r in range(50, 66):
                    nm = self._read(ws, f"AP{r}")
                    if nm:
                        names.append(f"{r}:{nm}")
                raise ExcelError(
                    f"{nav_cell} 的 {btn} 对应行类型是「{row_type}」而不是「{want}」，"
                    f"行号可能配错。当前面板：{' | '.join(names[:8])}")

            # 行类型校验通过，说明这就是该行的正确面板；记下来供下次（含新进程）识别
            self._panel_cache[self._panel_cache_key(nav_cell)] = self._panel_fingerprint()
            self._save_panel_cache()

            self.engine._retry(lambda: ws.Range(btn).Select(), attempts=5, delay=2)

            deadline = time.time() + (wait_sec or self.nav_cfg["open_wait_sec"])
            while time.time() < deadline:
                try:
                    last_names = self.engine.workbook_names()
                except Exception:
                    time.sleep(0.5)
                    continue
                new = [n for n in set(last_names) - before if "导航" not in n]
                if pattern:
                    new = [n for n in new if fnmatch.fnmatchcase(n, pattern)]
                elif expect == "package":
                    new = [n for n in new if "运行包" in n]
                if new:
                    log.info(f"{nav_cell} → 打开 {new[0]}")
                    return new[0]
                time.sleep(0.5)

            # 上一轮可能"晚成功"：先看看匹配的工作簿是不是已经在了
            late = self._find_opened(pattern, expect)
            if late:
                log.info(f"{nav_cell} → 打开 {late}（上一轮延迟完成）")
                return late
            log.warning(f"第 {attempt} 次点击 {btn} 未打开工作簿，重试（当前：{last_names}）")
            time.sleep(2)
        if via_cache:
            # 面板内容与上次验证过的正确指纹一致、按钮也是「打开」，点下去却毫无反应：
            # 典型是会话已死但面板停在旧内容。按会话失效处理，强制重启恢复。
            raise SessionExpiredError(
                f"选中 {nav_cell} 并点击 {btn} 后未打开新工作簿（面板内容与上次一致，"
                "按钮为「打开」）。AO/BPC 会话可能已过期，需要重新打开 Excel 并登录。")
        raise ExcelError(f"选中 {nav_cell} 并点击 {btn} 后未打开新工作簿（已重试 3 次）")

    def _find_opened(self, pattern: Optional[str], expect: str) -> Optional[str]:
        """当前已打开的工作簿里，是否已有符合期望的那个（用于重试兜底）。"""
        if not pattern and expect != "package":
            return None                    # expect='any' 无法区分，交给调用方自己查
        try:
            names = [n for n in self.engine.workbook_names() if "导航" not in n]
        except Exception:
            return None
        for name in names:
            if pattern and fnmatch.fnmatchcase(name, pattern):
                return name
            if not pattern and "运行包" in name:
                return name
        return None

    # ---------- 填表 ----------

    def fill_form(self, step: dict) -> dict:
        raise_if_cancelled()               # 安全点：还没打开表单就取消，零副作用
        form_id = step["form"]
        cfg = self.form_cfgs[form_id]
        pattern = cfg["workbook_pattern"]

        self._close_if_open(pattern)
        self.open_via_nav(step["nav_cell"], self.nav_cfg["open_button_cell"], "form",
                          pattern=pattern, expect_labels=step.get("nav_expect"))
        self.engine.get_workbook(pattern, wait_sec=30)

        # 锁定表单不要尝试切版本/写数据；命中后由 run_flow 跳过并结束流程
        lock = FormWriter(self.engine, cfg).detect_lock()
        if lock:
            raise FormLockedError(form_id, str(cfg.get("label") or form_id), lock)

        # 从月初版快捷端口打开的表单可能仍是月初版：逐表核对并刷新到目标版本
        self.ensure_form_version(form_id)
        self.engine.get_workbook(pattern, wait_sec=30)

        data = self._load_data(self.data_dir / step["data"])
        writer = FormWriter(self.engine, cfg)
        pf = writer.preflight(data)
        if not pf.ok:
            failed = [f"{name}: {detail}" for name, ok, detail in pf.checks if not ok]
            if pf.skipped:
                failed.append(f"未匹配科目: {pf.skipped[:5]}")
            raise ExcelError(f"{form_id} preflight 未通过 — " + "；".join(failed or ["无命中"]))
        log.info(f"{form_id}: 命中 {len(pf.matched_rows)} 格")
        # 安全点：preflight 完成后、真正写入 BPC 之前
        raise_if_cancelled()
        res = writer.write_and_save(pf, verify_after=False)
        if not res.get("saved"):
            return res

        # 保存"没报错"不等于数据真的进了 BPC：关闭不保存 → 重开 → 读回本次写入的格子。
        # 任一格对不上就判该步骤失败，避免带着错误数据继续跑计算和运行包。
        enabled = bool(self.wb_cfg.get("enabled", True))
        if cfg.get("verify_writeback") is False:
            enabled = False
        if enabled:
            fails = self.verify_persisted(form_id, step["nav_cell"], res.get("cells") or {},
                                          expect_labels=step.get("nav_expect"))
            res["verified"] = not fails
            if fails:
                res.setdefault("failures", []).extend(f"落库校验 {f}" for f in fails[:5])
                if len(fails) > 5:
                    res["failures"].append(f"落库校验失败共 {len(fails)} 格（仅列出前 5 格）")
        return res

    def verify_persisted(self, form_id: str, nav_cell: str, cells: dict,
                         expect_labels: Optional[list[str]] = None) -> list[str]:
        """关闭工作簿（不保存）→ 经导航表重开 → 读回比对，确认数据已落到 BPC。

        返回不一致项列表（空列表 = 全部通过）。可用 flow.yaml 的
        `verify_writeback.sample` 抽样加速；默认全量，一次 600 格的读回约几秒。
        """
        raise_if_cancelled()               # 安全点：数据已保存，取消后不再读回
        cfg = self.form_cfgs[form_id]
        pattern = cfg["workbook_pattern"]
        items = list(cells.items())
        sample = int(self.wb_cfg.get("sample") or 0)
        if sample and len(items) > sample:
            step = max(1, len(items) // sample)
            items = items[::step][:sample]
        if not items:
            return []
        tol = float(self.wb_cfg.get("tolerance", 0.01))

        self._close_if_open(pattern)
        self.open_via_nav(nav_cell, self.nav_cfg["open_button_cell"], "form",
                          pattern=pattern, expect_labels=expect_labels)
        self.ensure_form_version(form_id)
        wb = self.engine.get_workbook(pattern, wait_sec=60)
        ws = self.engine.open_sheet(wb, cfg["sheet"])

        failures: list[str] = []
        for n, (cell, expected) in enumerate(items):
            if n and n % 100 == 0:
                raise_if_cancelled()       # 全量读回时也给取消留检查点
            actual = self._read(ws, cell)
            if _blank(expected) and _blank(actual):
                continue
            try:
                ok = abs(float(actual or 0) - float(expected or 0)) <= tol
            except (TypeError, ValueError):
                ok = str(actual).strip() == str(expected).strip()
            if not ok:
                failures.append(f"{cell}: 期望 {expected!r}, 重开后 {actual!r}")

        # 故意不关闭：后续可能需要在这张表上执行计算宏（如 Execute_PS），
        # 由调用方在收尾时统一 close_form。
        log.info(f"{form_id}: 落库校验 {len(items)} 格，失败 {len(failures)} 格")
        return failures

    def preflight_form(self, step: dict) -> dict:
        """打开编制表 → 只做预检（不写任何格子）→ 关闭。用于 --preflight-only 干跑。"""
        form_id = step["form"]
        cfg = self.form_cfgs[form_id]
        pattern = cfg["workbook_pattern"]

        self._close_if_open(pattern)
        self.open_via_nav(step["nav_cell"], self.nav_cfg["open_button_cell"], "form",
                          pattern=pattern, expect_labels=step.get("nav_expect"))
        self.engine.get_workbook(pattern, wait_sec=30)

        lock = FormWriter(self.engine, cfg).detect_lock()
        if lock:
            self._close_if_open(pattern)
            return {
                "form": form_id, "ok": False, "matched": 0, "skipped": [],
                "checks": [("锁定", False, lock)], "locked": lock,
            }

        self.ensure_form_version(form_id)
        self.engine.get_workbook(pattern, wait_sec=30)

        data = self._load_data(self.data_dir / step["data"])
        pf = FormWriter(self.engine, cfg).preflight(data)
        out = {
            "form": form_id,
            "ok": pf.ok,
            "matched": len(pf.matched_rows),
            "skipped": pf.skipped,
            "checks": [(label, bool(pass_), detail) for label, pass_, detail in pf.checks],
        }
        self._close_if_open(pattern)
        return out

    def close_form(self, form_id: str) -> None:
        self._close_if_open(self.form_cfgs[form_id]["workbook_pattern"])

    # ---------- 计算 / 运行包 ----------

    def trigger_calc(self, form_id: str, macro: str) -> None:
        cfg = self.form_cfgs[form_id]
        wb = self.engine.get_workbook(cfg["workbook_pattern"], wait_sec=20)
        t0 = time.time()
        self.engine.run_macro(macro, wb)
        log.info(f"{form_id}: 执行 {macro} 完成（{int((time.time()-t0)*1000)}ms）")
        time.sleep(3)

    def run_package(self, nav_cell: str, step_row: int, expected_name: str = "",
                    expect_labels: Optional[list[str]] = None) -> dict:
        """打开运行包清单 → Execute → 轮询状态 → 关闭。

        expected_name：该行号期望的包名（flow.yaml after_close.packages 的
        「行号: 包名」写法）。步骤面板若停在别的同类型行上，会开到错误的
        包上下文——清单 AB73 的步骤名与期望对不上就中止，绝不执行错包。
        expect_labels：导航行 AA+AB+AC 必须包含的关键字；None = 单元级默认。
        """
        self.ensure_navigator_version()
        self._close_if_open(self.pkg_cfg["workbook_pattern"])
        self.open_via_nav(nav_cell, open_cell=f"AT{step_row}",
                          expect="package", wait_sec=20,
                          expect_labels=expect_labels)

        wb = self.engine.get_workbook(self.pkg_cfg["workbook_pattern"], wait_sec=30)
        ws = self.engine.open_sheet(wb, self.pkg_cfg["sheet"])

        version = str(self._read(ws, self.pkg_cfg.get("version_cell", "AB60")) or "")
        step_name = str(self._read(ws, self.pkg_cfg["step_cell"]) or "")
        expected = str(expected_name or "").strip()
        if expected and expected not in step_name:
            try:
                self.engine.close_workbook(wb, save=False)
            except Exception as exc:
                log.warning(f"关闭不符的运行包清单失败（忽略）: {exc}")
            raise ExcelError(
                f"运行包步骤名不符：清单显示「{step_name}」，期望「{expected}」。"
                "可能是步骤面板停在了别的同类型行上（不同行的按钮位置相同），"
                "或 after_close.packages 的行号/包名配置有误；"
                "已关闭清单、未执行任何宏，请核对后重试")
        log.info(f"运行包: {step_name} | 版本: {version}")

        raise_if_cancelled()               # 安全点：还没触发运行包，取消无副作用
        self.engine.run_macro(self.pkg_cfg["execute_macro"], wb)
        log.info("已触发执行，等待完成…")

        status, duration = "", ""
        pattern = self.pkg_cfg["workbook_pattern"]
        sheet_name = self.pkg_cfg["sheet"]
        deadline = time.time() + self.pkg_cfg["max_wait_sec"]
        while time.time() < deadline:
            time.sleep(self.pkg_cfg["poll_interval_sec"])
            if is_cancelled():
                # 运行包已在后台执行：这里只是不再等它，不能假装它被撤销
                raise CancelledError(
                    f"已取消（运行包「{step_name}」已在后台触发，可能仍会执行完成）")
            try:
                # Refresh_Sheet 会重算/重建 AO 对象，旧的 wb/ws 可能失效；
                # 每次刷新前后都按名字重新绑定，避免读到断开的 COM 引用。
                wb = self.engine.get_workbook(pattern, wait_sec=15)
                ws = self.engine.open_sheet(wb, sheet_name)
                self.engine.run_macro(self.pkg_cfg["refresh_macro"], wb)
                wb = self.engine.get_workbook(pattern, wait_sec=15)
                ws = self.engine.open_sheet(wb, sheet_name)
            except ExcelError as exc:
                log.warning(f"刷新状态失败（继续轮询）: {exc}")
                continue
            time.sleep(2)
            status = str(self._read(ws, self.pkg_cfg["status_cell"]) or "")
            duration = str(self._read(ws, self.pkg_cfg["duration_cell"]) or "")
            log.info(f"状态={status} 时长={duration}")
            if "已完成" in status or "成功" in status:
                break
            if "失败" in status or "错误" in status:
                break

        self.engine.close_workbook(wb, save=False)
        ok = "已完成" in status or "成功" in status
        return {"package": step_name, "status": status, "duration": duration, "ok": ok}

    # ---------- 损益核对 ----------

    def read_indicators(self, allow_all_zero: bool = False) -> dict:
        """读损益指标。

        allow_all_zero=True 用于「清 0」流程：目标期间本来就被清成全 0，
        指标全 0 是预期结果，只告警不报错（否则清 0 会被误判为运行包没生效）。
        """
        raise_if_cancelled()
        self.indicators_all_zero = False
        vcfg = self.cfg["verify"]
        self.ensure_navigator_version()
        self._close_if_open("YC_CBG损益查询表*")
        self.open_via_nav(vcfg["nav_cell"], open_cell=f"AT{vcfg['form_row']}",
                          expect="any", wait_sec=20,
                          expect_labels=vcfg.get("nav_expect"))
        wb = self.engine.get_workbook("YC_CBG损益查询表*", wait_sec=30)
        ws = self.engine.open_sheet(wb, vcfg["sheet"])

        if vcfg.get("refresh_before_read"):
            time.sleep(2)
            self.engine.refresh(wb, macro_hint="Refresh_Sheet")
            time.sleep(12)                 # 等 AO 重算期间列（2026 全年 → 2027.01~04）
            # Refresh_Sheet 可能重建 AO 对象；重新按名字绑定工作簿/Sheet 再读表头
            try:
                wb = self.engine.get_workbook("YC_CBG损益查询表*", wait_sec=30)
                ws = self.engine.open_sheet(wb, vcfg["sheet"])
            except ExcelError as exc:
                log.warning(f"刷新后重新绑定损益表失败（沿用原引用）: {exc}")

        # 按表头动态解析期间列；若目标期间被前置年度列挤占，先折叠上一年度再读
        header_row = int(vcfg.get("header_row", 54))
        target = self.target_period
        actual = self._read_period_headers(ws, header_row)
        need = [target] if target else []
        need += [m for m in self._month_headers(actual) if target and m <= target]
        missing = [p for p in need if p not in actual]
        if missing:
            year = self._collapse_year_for(actual, target)
            if year:
                log.info(f"表头未找到 {missing}，折叠 {year} 年度列后重试")
                self._collapse_year_header(ws, header_row, year)
                actual = self._read_period_headers(ws, header_row)

        # 目标期间列必须存在：折叠失败/刷新未完成时不允许"能读多少读多少"，
        # 否则会静默输出 2026 年度合计或部分月份，看起来成功、实际答非所问。
        if target and target not in actual:
            found = ", ".join(sorted(actual)) or "(空)"
            raise ExcelError(
                f"损益表未出现目标期间列 {target}（当前表头期间: {found}）。"
                f"可能是年度列折叠失败、刷新未完成，或数据期间不在当前版本窗口内。"
                f"目标期间由数据目录名推导，请确认版本与期间匹配后重跑。")

        periods = self._select_periods(actual, target)
        if not periods:
            # 表头解析失败时退回配置里的固定列
            periods = {str(k): v for k, v in (vcfg.get("periods") or {}).items()}
            log.warning(f"表头未解析出期间列，改用配置兜底: {periods}")
            if target and target not in periods:
                raise ExcelError(
                    f"损益表头解析失败，且配置兜底期间里没有 {target}。"
                    f"请检查 verify.header_row / verify.periods 配置，或手动确认损益表状态。")
        log.info(f"期间列: {periods}")

        out: dict[str, dict[str, Any]] = {}
        for name, row in vcfg["indicator_rows"].items():
            out[name] = {}
            for period, col in periods.items():
                out[name][period] = self._read(ws, f"{col}{row}")

        # 指标合理性：目标期间全空/全 0 基本意味着运行包没生效，不能当成功结果返回
        if target:
            self._check_target_indicators(out, target, allow_all_zero)
        return out

    def _check_target_indicators(self, out: dict, target: str,
                                 allow_all_zero: bool) -> bool:
        """校验目标期间指标不全为空/0；返回是否全 0。

        全 0 基本意味着运行包没生效，不能当成功结果返回——除非是清 0 流程
        （allow_all_zero=True），此时全 0 正是目标状态。
        """
        vals = {name: _num(rows.get(target)) for name, rows in out.items()}
        if all(v is None or v == 0 for v in vals.values()):
            if not allow_all_zero:
                raise ExcelError(
                    f"损益核对 {target} 的指标全部为空或为 0，运行包可能未生效。"
                    f"请先确认三个运行包状态为「已完成」后重试。指标: {list(out)}")
            self.indicators_all_zero = True
            log.warning(f"损益核对 {target} 指标全部为 0（清 0 模式，按预期放行）")
            return True
        zero = [name for name, v in vals.items() if v is None or v == 0]
        if zero:
            log.warning(f"损益核对 {target} 以下指标为 0/空: {zero}")
        return False

    # ---------- 期间解析 ----------

    @staticmethod
    def _month_headers(headers: dict[str, str]) -> list[str]:
        """表头里所有 'YYYY.MM' 形式的期间，升序。"""
        return sorted(k for k in headers if re.fullmatch(r"\d{4}\.\d{2}", str(k).strip()))

    @staticmethod
    def _select_periods(headers: dict[str, str], target: str) -> dict[str, str]:
        """按表头挑要读的列：先年度合计（YYYY），再截至目标期间的月份（YYYY.MM）。"""
        years = sorted(k for k in headers if re.fullmatch(r"\d{4}", str(k).strip()))
        months = [m for m in PbfFlow._month_headers(headers) if not target or m <= target]
        out = {y: headers[y] for y in years}
        out.update({m: headers[m] for m in months})
        return out

    def _collapse_year_for(self, headers: dict[str, str], target: str) -> str:
        """要折叠的年度：表头里比目标期间小的最大年度。

        例：目标 2027.03、表头有 2026 → 折叠 2026；目标 2026.12 → 折叠 2025（通常不存在，等于不折叠）。
        """
        if not target:
            return str(self.cfg["verify"].get("collapse_year") or "")
        years = sorted(y for y in headers
                       if re.fullmatch(r"\d{4}", str(y).strip()) and y < target[:4])
        if years:
            return years[-1]
        try:
            return str(int(target[:4]) - 1)
        except ValueError:
            return ""

    def _collapse_year_header(self, ws, header_row: int, year: str) -> bool:
        """点击年度表头（如 AB54 的 2026）折叠该年度，让后面的年度期间列显示出来。

        AO 把展开状态写在表头单元格的数字格式里：'[+] ' = 已折叠，'[-] ' = 已展开。
        表头文字是居中的，折叠标记不一定在单元格左边缘，所以从中间向两侧试几个点位；
        每次点击后要等 AO 重算完，状态没变才换下一个点位。
        """
        from .scanner import col_letter

        addr = None
        for i in range(1, 151):
            letter = col_letter(i)
            try:
                v = self._read(ws, f"{letter}{header_row}")
            except Exception:
                continue
            if v is not None and str(v).strip() == year:
                addr = f"{letter}{header_row}"
                break
        if addr is None:
            log.warning(f"表头第 {header_row} 行未找到年度 {year}")
            return False

        def collapsed() -> bool:
            # 格式码形如 '"[+] "@'（注意首字符是格式里的引号），所以只能判断子串
            try:
                return "[+]" in str(ws.Range(addr).NumberFormat)
            except Exception:
                return False

        if collapsed():
            log.info(f"{addr}={year} 已是折叠状态")
            return True

        for frac in (0.5, 0.45, 0.55, 0.4, 0.6, 0.35, 0.65):
            self.engine.click_cell(ws, addr, x_frac=frac)
            for _ in range(8):              # 最多等 8×1.5s 让 AO 重算完
                time.sleep(1.5)
                if collapsed():
                    log.info(f"点击 {addr}({frac:.2f}) 后 {year} 已折叠")
                    return True
            log.info(f"点击 {addr}({frac:.2f}) 未生效，换下一个点位")
        log.warning(f"{addr}={year} 折叠失败")
        return False

    def _read_period_headers(self, ws, header_row: int, col_from: int = 1, col_to: int = 150) -> dict[str, str]:
        """读取表头行，返回 {期间文本: 列字母}。扫描范围要足够宽（2027 列可能被 2026 挤到很右侧）。"""
        from .scanner import col_letter
        found: dict[str, str] = {}
        prev: Optional[tuple[int, int]] = None
        for i in range(col_from, col_to + 1):
            letter = col_letter(i)
            try:
                v = self._read(ws, f"{letter}{header_row}")
            except Exception:
                continue
            if v is None:
                continue
            # AO 表头可能是文本 / 数字 / 日期，统一归一化后再匹配。
            # 单个异常单元格（NaN/inf、异常公式）不能让整轮核对中断；
            # 只捕获归一化里可能抛的转换类异常（float/int 对怪值），别把
            # COM/属性错误也吞掉。
            try:
                key, new_prev = normalize_period_header(v, prev)
            except (ValueError, TypeError, OverflowError) as exc:
                log.warning(f"表头 {letter}{header_row}={v!r} 解析失败，已跳过: {exc}")
                continue
            if new_prev != prev:
                prev = new_prev
            # 只保留期间键（'YYYY' / 'YYYY.MM'）：NaN 经 str() 会变成 "nan"，
            # 杂散文本也会进表，混进 found 会污染"未找到期间"的报错信息。
            if key and key not in found and re.fullmatch(r"\d{4}(\.\d{2})?", key):
                found[key] = letter
        return found


# ---------- CLI 入口 ----------

def close_form_workbooks_on_cancel(engine: ExcelEngine, form_cfgs: dict,
                                   flow_cfg: dict) -> list[str]:
    """取消收尾：关闭本工具打开的表单/运行包工作簿（不保存），返回关闭的名字。

    只匹配配置里的 workbook_pattern：导航表和用户自己打开的工作簿一律不碰。
    关闭时丢弃本地改动——取消点若在 Plan_Save 之前，BPC 里本来就没有这些值；
    若在之后，数据已经保存，丢弃本地副本也没有影响。
    """
    patterns: list[str] = []
    for cfg in form_cfgs.values():
        pat = cfg.get("workbook_pattern")
        if pat:
            patterns.append(str(pat))
    pkg = flow_cfg.get("package") or {}
    if pkg.get("workbook_pattern"):
        patterns.append(str(pkg["workbook_pattern"]))
    if not patterns or engine is None:
        return []
    try:
        names = engine.workbook_names()
    except Exception as exc:
        log.warning(f"取消收尾：无法枚举工作簿（{exc}），请确认 Excel 中没有未保存的表单")
        return []
    closed: list[str] = []
    for name in names:
        if "导航" in name:
            continue
        if not any(fnmatch.fnmatchcase(name, pat) for pat in patterns):
            continue
        try:
            wb = engine.find_workbook(name)
            engine.close_workbook(wb, save=False)
            closed.append(name)
        except Exception as exc:
            log.warning(f"取消收尾：关闭 {name} 失败（忽略）: {exc}")
    if closed:
        log.warning(f"已按取消请求关闭未完成的工作簿（不保存）：{closed}")
    return closed


def await_background_login(engine: ExcelEngine, flow_cfg: dict,
                           wait_sec: float) -> bool:
    """分片等待启动阶段的后台登录线程出结果；确认成功后补 AO 收尾屏障。

    之前是单次 handle.wait(timeout + 30)：最长阻塞 150s 且中途不检查取消
    标志，Web 端取消只能干等宽限期后被硬杀进程树。改为按秒分片，取消一
    出现就抛 CancelledError 走安全点收尾。

    登录确认成功且线程真的填过登录框时，AO 加载项会自动刷新数据源（实测
    占住 UI 线程 30s 级），必须等它收尾才能继续调 SAP 变量——复用 Excel 的
    路径在 bootstrap 里只起线程不等结果，屏障就补在这里。login_settled
    防止 bootstrap 已做过的屏障重复执行。
    """
    handle = getattr(engine, "login_handle", None)
    if handle is not None and not handle.finished:
        deadline = time.time() + max(wait_sec, 0.0)
        while True:
            raise_if_cancelled()
            if handle.wait(timeout=1) is not None:
                break
            if time.time() >= deadline:
                break
    confirmed = bool(handle is not None and handle.finished
                     and bool(handle.outcome))
    if confirmed and getattr(handle, "saw_box", False) \
            and not getattr(engine, "login_settled", False):
        wait_ao_login_idle(engine, flow_cfg)
    return confirmed


def run_flow(engine: ExcelEngine, flow_cfg: dict, form_cfgs: dict, data_dir: Path,
             skip_fill: bool = False, only_verify: bool = False,
             expected_version: Optional[str] = None,
             from_step: Optional[str] = None,
             preflight_only: bool = False,
             switch_version_key: Optional[str] = None,
             clear_mode: bool = False) -> dict:
    """from_step: 从第几步开始（1 开始的序号，或表单 ID），之前的步骤整步跳过。

    跳过的步骤连它的计算宏和运行包一起跳过（这些属于该步骤的收尾动作）。
    clear_mode: 清 0 流程。目标期间损益指标全 0 属预期结果，核对只告警不报错。
    """
    # 面板指纹跨进程落盘：上次运行可能把面板停在某一行，新进程要能识别出来，
    # 否则"选中前后内容一致"会被误判为会话失效（见 logs/nav_panel_cache.json）
    flow = PbfFlow(engine, flow_cfg, form_cfgs, data_dir, persist_panel_cache=True)
    report: dict[str, Any] = {"version": "", "steps": [], "packages": [], "indicators": {}}

    steps = flow_cfg["steps"]
    start_idx = 0
    if from_step:
        key = str(from_step).strip()
        if key.isdigit():
            start_idx = int(key) - 1
            if not 0 <= start_idx < len(steps):
                raise ExcelError(f"--from-step {key} 超出范围（共 {len(steps)} 步）")
        else:
            idx = next((i for i, s in enumerate(steps) if s["form"] == key), None)
            if idx is None:
                names = [s["form"] for s in steps]
                raise ExcelError(f"--from-step 未找到表单 '{key}'，可选：{names}")
            start_idx = idx
        report["from_step"] = steps[start_idx]["form"]

    # 0. 登录（检测到登录框则自动填写；启动阶段已处理过则跳过）
    login_cfg = login_config(flow_cfg)
    if getattr(engine, "login_done", False):
        log.info("启动阶段已处理 AO 登录，跳过重复登录")
        report["login"] = True
    elif login_cfg.get("username"):
        # 启动阶段可能起过后台登录线程：以句柄里的"登录成功信号"为准，
        # 不能拿"线程已启动/已结束"当成功——密码错误时线程照样退出，
        # 若被当成已登录，后续 COM 调用会被模态登录框永久阻塞
        handle = getattr(engine, "login_handle", None)
        if await_background_login(engine, flow_cfg,
                                  float(login_cfg.get("timeout", 60)) + 30):
            log.info("启动阶段后台登录已成功，跳过重复登录")
            report["login"] = True
        else:
            if handle is not None and handle.finished and not handle.outcome \
                    and handle.saw_box:
                # 见过登录框但没登录成功：密码错误或登录框卡住，
                # 继续跑只会把清晰的报错变成 COM 挂死
                raise ExcelError(
                    f"AO 自动登录失败（{handle.error or '超时或密码错误'}），"
                    "登录框可能仍在——为避免账号锁定，工具不会自动重试；"
                    "请在 Excel 中手动登录后重试")
            from .ao_login import AoLogin
            launched = bool(getattr(engine, "launched_now", False))
            login = AoLogin(
                username=login_cfg["username"],
                password=login_cfg.get("password", ""),
                client=str(login_cfg.get("client", "333")),
                lang=str(login_cfg.get("lang", "zh")),
            )
            ok = login.ensure(float(login_cfg.get("timeout", 60)),
                              # 刚由本工具启动 Excel 时，登录框还没弹出来，必须等它出现；
                              # 但超过登录宽限仍没有框时按缓存会话继续，不再空等到超时
                              expect_login=launched,
                              no_box_ok_after=launch_no_box_ok_sec(flow_cfg) if launched else None)
            if not ok:
                raise ExcelError(
                    "AO 自动登录失败（已停止自动重试，避免账号锁定），"
                    "请在 Excel 中手动登录后重试")
            if getattr(login, "saw_login_box", False) \
                    and not getattr(engine, "login_settled", False):
                # 同步登录真的填过登录框：加载项随后会自动刷新数据源，
                # 等它收尾再继续（同 wait_ao_login_idle 的屏障）
                wait_ao_login_idle(engine, flow_cfg)
            report["login"] = True

    rcfg = flow_cfg.get("recovery") or {}
    recovery_on = bool(rcfg.get("enabled", True))
    max_retries = int(rcfg.get("max_retries", 1)) if recovery_on else 0

    def run_with_recovery(fn, label: str):
        """执行 fn；若因 Excel 断开失败且开启了恢复，重开 Excel 后重做一次。"""
        attempts = max_retries + 1
        for attempt in range(1, attempts + 1):
            steps_len, pkgs_len = len(report["steps"]), len(report["packages"])
            try:
                return fn()
            except Exception as exc:            # 包含未包装的 com_error
                if isinstance(exc, CancelledError):
                    # 取消不是故障：不重试、不走恢复，直接把工作簿留在
                    # "未保存"状态交给上层收尾
                    raise
                del report["steps"][steps_len:]      # 丢掉这次重试前的半成品记录
                del report["packages"][pkgs_len:]
                # 会话失效有两个来源：已翻译成 SessionExpiredError 的，和
                # trigger_calc/run_package 里宏直接把 0x800A9C68 抛上来的
                # （不走 refresh_flow 的翻译，这里按文本特征兜底识别）。
                session_expired = isinstance(exc, SessionExpiredError) or is_session_expired(exc)
                if attempt >= attempts or not (is_disconnect(exc) or session_expired):
                    raise
                reason = "AO 会话失效" if session_expired else "Excel 断开"
                log.warning(f"{label} 因{reason}失败：{exc}")
                if not flow.recover(session_expired=session_expired):
                    raise
                log.info(f"Excel 已恢复，重做：{label}")

    # 5) 切版本（用户路径第 ⑤ 步）：先切到选定版本，再「刷新状态」。
    dims_file = app_root() / "config" / "dimensions.yaml"
    dims = yaml.safe_load(dims_file.read_text(encoding="utf-8")) or {}
    dims_nav = dims.get("navigator") or {}
    versions = dims.get("versions") or {}
    # 版本单元格在 dimensions.yaml，其余导航配置在 flow.yaml；合并给 Navigator 用
    flow.nav_dims_cfg = {**dims_nav, **flow.nav_cfg}
    switched_display: Optional[str] = None
    if switch_version_key:
        switched_display = versions.get(switch_version_key)
        run_with_recovery(
            lambda: Navigator(engine, dims_nav).switch_version(
                switch_version_key, switched_display),
            f"切换版本 {switch_version_key}")
        report["switched_version"] = switch_version_key

    expected = expected_version or switched_display
    report["version"] = run_with_recovery(
        lambda: flow.confirm_version(expected), "版本确认")
    flow.check_version_window(report["version"])

    # 表单级版本：每张表打开后按这个目标核对，不是就切并刷新。
    # 优先用 --version-key；否则按显示名反查；再退回导航表 C20 当前成员。
    form_key = switch_version_key
    if not form_key and expected_version:
        form_key = next(
            (k for k, d in versions.items() if _version_matches(expected_version, d)), None)
    if not form_key:
        try:
            form_key = Navigator(engine, dims_nav).current_version_key()
        except ExcelError as exc:
            log.warning(f"读取导航表当前版本成员失败，表单将只核对显示文本: {exc}")
    if form_key:
        flow.form_version_key = form_key
        flow.form_version_display = versions.get(form_key) or report["version"] or expected_version
        report["form_version_key"] = form_key
        report["form_version_display"] = flow.form_version_display
        log.info(f"表单目标版本: {flow.form_version_display} ({form_key})")
    # AO 刚登录时 Refresh_Status 偶发瞬时失败，内部重试；仍失败才交给恢复机制
    run_with_recovery(flow.refresh_flow_retry, "刷新流程")

    if preflight_only:
        # 干跑：逐张表打开→匹配行→检查期间覆盖→关闭，全程不写入。
        report["preflights"] = []
        for idx, step in enumerate(steps[start_idx:], start_idx + 1):
            res = run_with_recovery(
                lambda s=step: flow.preflight_form(s), f"预检 {idx}（{step['form']}）")
            res["index"] = idx
            report["preflights"].append(res)
        return report

    def run_step(step: dict) -> None:
        form_id = step["form"]
        if not skip_fill:
            res = flow.fill_form(step)
            report["steps"].append({"form": form_id, "written": res.get("written"),
                                    "saved": res.get("saved"),
                                    "verified": res.get("verified"),
                                    "unchanged": res.get("unchanged", False),
                                    "failures": res.get("failures") or []})
            if not res.get("saved"):
                # 没保存成功就继续跑计算/运行包会把脏数据一路带下去，这里直接中断
                raise ExcelError(
                    f"{form_id} 保存失败：{'; '.join(res.get('failures') or ['原因未知'])}")
            if res.get("verified") is False:
                # 保存没报错但重开后读不回写入值：同样不能继续，否则计算基于旧数据
                raise ExcelError(
                    f"{form_id} 落库校验失败："
                    + "; ".join((res.get("failures") or ["重开后读回的数值与写入值不一致"])[:3])
                    + "。请确认表格是否进入编制态、数值是否被 AO 自动重算；"
                      "确认前不要继续后续步骤。")

        after = step.get("after_save") or {}
        if after.get("calc_macro"):
            flow.trigger_calc(form_id, after["calc_macro"])
            report["steps"].append({"form": form_id, "calc": after["calc_macro"]})

        if not skip_fill:
            flow.close_form(form_id)

        after_close = step.get("after_close") or {}
        if after_close:
            for row, expected in _package_items(after_close.get("packages")):
                r = flow.run_package(after_close["nav_cell"], row,
                                     expected_name=expected,
                                     expect_labels=after_close.get("nav_expect"))
                report["packages"].append(r)
                if not r["ok"]:
                    raise ExcelError(f"运行包未成功: {r}")

    if only_verify:
        report["indicators"] = run_with_recovery(
            lambda: flow.read_indicators(allow_all_zero=clear_mode), "损益核对")
        report["indicators_all_zero"] = flow.indicators_all_zero
        return report

    for i, step in enumerate(steps):
        raise_if_cancelled()               # 安全点：取消后不再开始新表单
        form_id = step["form"]
        if i < start_idx:
            log.info(f"跳过步骤 {i + 1}（{form_id}）")
            continue
        try:
            run_with_recovery(lambda s=step: run_step(s), f"步骤 {i + 1}（{form_id}）")
        except FormLockedError as exc:
            # 业务要求：变动费用等表单被其他用户锁定时跳过该表并结束流程，
            # 不要把后面的计算/运行包/损益核对带着旧数据继续跑。
            report["locked"] = {
                "form": exc.form_id, "label": exc.label,
                "user": exc.user, "message": str(exc),
            }
            # 是否因锁定中止，看 report["locked"] 是否存在即可（CLI 消费它），
            # 不再另设 report["stopped"] 冗余标记
            log.error(str(exc))
            return report

    report["indicators"] = run_with_recovery(
        lambda: flow.read_indicators(allow_all_zero=clear_mode), "损益核对")
    report["indicators_all_zero"] = flow.indicators_all_zero
    return report
