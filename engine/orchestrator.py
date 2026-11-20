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
import logging
import os
import re
import time
from pathlib import Path
from typing import Any, Optional

import pandas as pd
import yaml

from .excel import ExcelEngine, ExcelError, is_disconnect
from .bootstrap import ensure_excel_ready
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


class PbfFlow:
    def __init__(self, engine: ExcelEngine, flow_cfg: dict, form_cfgs: dict, data_dir: Path) -> None:
        self.engine = engine
        self.cfg = flow_cfg
        self.nav_cfg = flow_cfg["navigator"]
        self.pkg_cfg = flow_cfg["package"]
        self.wb_cfg = flow_cfg.get("verify_writeback") or {}
        self.form_cfgs = form_cfgs
        self.data_dir = Path(data_dir)

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
        if expected and expected not in v:
            raise ExcelError(f"版本不符：当前 '{v}'，期望包含 '{expected}'")
        log.info(f"版本确认: {v}")
        return v

    def refresh_flow(self) -> None:
        wb = self._nav_wb()
        try:
            self.engine.run_macro(self.nav_cfg["refresh_flow_macro"], wb)
        except ExcelError as exc:
            text = str(exc)
            # "发生意外"（0x80020009 / 内层 0x800A9C68）几乎都来自 AO/BPC 服务端：
            # 会话过期，或版本与目标期间不匹配。给一句能直接下手的人话。
            if ("发生意外" in text or "-2147352567" in text
                    or "-2146788248" in text or "0x800A9C68" in text):
                raise ExcelError(
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

    def recover(self) -> bool:
        """Excel 断开后自动恢复：重开导航表 → 重新登录 → 换上新的引擎。

        返回 True 表示已恢复，调用方可以重做失败的步骤；False 表示恢复失败。
        """
        rcfg = self.cfg.get("recovery") or {}
        log.warning("Excel 连接已断开，开始自动恢复…")
        # 只结束本流程连接的那个 Excel 进程；不碰用户同时打开的其它 Excel。
        stale_pid = getattr(self.engine, "pid", None)
        try:
            self.engine = ensure_excel_ready(
                self.cfg, kill_stale=bool(rcfg.get("kill_stale_excel", True)),
                stale_pid=stale_pid)
        except Exception as exc:
            log.error(f"自动恢复失败：{exc}")
            return False

        login_cfg = login_config(self.cfg)
        # 启动阶段的后台登录线程以结果句柄为准：成功才跳过登录，
        # 失败/未确认都要在这里重新同步登录，不能把"线程已结束"当已登录
        handle = getattr(self.engine, "login_handle", None)
        if handle is not None and not handle.finished:
            handle.wait(float(login_cfg.get("timeout", 120)) + 30)
        login_confirmed = bool(getattr(self.engine, "login_done", False)) or (
            handle is not None and handle.finished and bool(handle.outcome))
        if login_cfg.get("username") and not login_confirmed:
            from .ao_login import AoLogin

            ok = AoLogin(
                username=login_cfg["username"],
                password=login_cfg.get("password", ""),
                client=str(login_cfg.get("client", "333")),
                lang=str(login_cfg.get("lang", "zh")),
            ).ensure(float(login_cfg.get("timeout", 120)),
                     expect_login=bool(getattr(self.engine, "launched_now", False)))
            if not ok:
                log.error("恢复后自动登录失败")
                return False
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

    def open_via_nav(self, nav_cell: str, open_cell: Optional[str] = None,
                     expect: str = "form", wait_sec: Optional[int] = None,
                     pattern: Optional[str] = None) -> str:
        """选中业务单元行 → 点"打开" → 返回新打开的工作簿名。

        expect: 'form' 期望打开编制表；'package' 期望打开运行包清单。
        pattern: 编制表的 workbook_pattern（有些表单名里没有"编制"二字，只能按模式匹配）。
        """
        wb = self._nav_wb()
        ws = self.engine.open_sheet(wb, self.nav_cfg["sheet"])
        btn = open_cell or self.nav_cfg["open_button_cell"]

        # "打开"靠导航表的 SelectionChange 触发，偶尔会被 Excel/AO 忙吞掉，因此重试几轮
        last_names: list[str] = []
        for attempt in range(1, 4):
            self.engine._retry(lambda: (wb.Activate(), ws.Activate()), attempts=5, delay=2)
            before = set(self.engine._retry(
                lambda: self.engine.workbook_names(), attempts=5, delay=2))
            self.engine._retry(lambda: ws.Range(nav_cell).Select(), attempts=5, delay=2)
            time.sleep(5)                  # 等右侧面板加载出"打开"
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
        form_id = step["form"]
        cfg = self.form_cfgs[form_id]
        pattern = cfg["workbook_pattern"]

        self._close_if_open(pattern)
        self.open_via_nav(step["nav_cell"], self.nav_cfg["open_button_cell"], "form",
                          pattern=pattern)
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
        res = writer.write_and_save(pf, verify_after=False)
        if not res.get("saved"):
            return res

        # 保存"没报错"不等于数据真的进了 BPC：关闭不保存 → 重开 → 读回本次写入的格子。
        # 任一格对不上就判该步骤失败，避免带着错误数据继续跑计算和运行包。
        enabled = bool(self.wb_cfg.get("enabled", True))
        if cfg.get("verify_writeback") is False:
            enabled = False
        if enabled:
            fails = self.verify_persisted(form_id, step["nav_cell"], res.get("cells") or {})
            res["verified"] = not fails
            if fails:
                res.setdefault("failures", []).extend(f"落库校验 {f}" for f in fails[:5])
                if len(fails) > 5:
                    res["failures"].append(f"落库校验失败共 {len(fails)} 格（仅列出前 5 格）")
        return res

    def verify_persisted(self, form_id: str, nav_cell: str, cells: dict) -> list[str]:
        """关闭工作簿（不保存）→ 经导航表重开 → 读回比对，确认数据已落到 BPC。

        返回不一致项列表（空列表 = 全部通过）。可用 flow.yaml 的
        `verify_writeback.sample` 抽样加速；默认全量，一次 600 格的读回约几秒。
        """
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
                          pattern=pattern)
        wb = self.engine.get_workbook(pattern, wait_sec=60)
        ws = self.engine.open_sheet(wb, cfg["sheet"])

        failures: list[str] = []
        for cell, expected in items:
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
                          pattern=pattern)
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

    def run_package(self, nav_cell: str, step_row: int) -> dict:
        """打开运行包清单 → Execute → 轮询状态 → 关闭。"""
        self._close_if_open(self.pkg_cfg["workbook_pattern"])
        self.open_via_nav(nav_cell, open_cell=f"AT{step_row}",
                          expect="package", wait_sec=20)

        wb = self.engine.get_workbook(self.pkg_cfg["workbook_pattern"], wait_sec=30)
        ws = self.engine.open_sheet(wb, self.pkg_cfg["sheet"])

        version = str(self._read(ws, self.pkg_cfg.get("version_cell", "AB60")) or "")
        step_name = str(self._read(ws, self.pkg_cfg["step_cell"]) or "")
        log.info(f"运行包: {step_name} | 版本: {version}")

        self.engine.run_macro(self.pkg_cfg["execute_macro"], wb)
        log.info("已触发执行，等待完成…")

        status, duration = "", ""
        pattern = self.pkg_cfg["workbook_pattern"]
        sheet_name = self.pkg_cfg["sheet"]
        deadline = time.time() + self.pkg_cfg["max_wait_sec"]
        while time.time() < deadline:
            time.sleep(self.pkg_cfg["poll_interval_sec"])
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

    def read_indicators(self) -> dict:
        vcfg = self.cfg["verify"]
        self._close_if_open("YC_CBG损益查询表*")
        self.open_via_nav(vcfg["nav_cell"], open_cell=f"AT{vcfg['form_row']}",
                          expect="any", wait_sec=20)
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
            vals = {name: _num(rows.get(target)) for name, rows in out.items()}
            if all(v is None or v == 0 for v in vals.values()):
                raise ExcelError(
                    f"损益核对 {target} 的指标全部为空或为 0，运行包可能未生效。"
                    f"请先确认三个运行包状态为「已完成」后重试。指标: {list(out)}")
            zero = [name for name, v in vals.items() if v is None or v == 0]
            if zero:
                log.warning(f"损益核对 {target} 以下指标为 0/空: {zero}")
        return out

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
            # 单个异常单元格（NaN/inf、异常公式）不能让整轮核对中断。
            try:
                key, new_prev = normalize_period_header(v, prev)
            except Exception as exc:
                log.warning(f"表头 {letter}{header_row}={v!r} 解析失败，已跳过: {exc}")
                continue
            if new_prev != prev:
                prev = new_prev
            if key and key not in found:
                found[key] = letter
        return found


# ---------- CLI 入口 ----------

def run_flow(engine: ExcelEngine, flow_cfg: dict, form_cfgs: dict, data_dir: Path,
             skip_fill: bool = False, only_verify: bool = False,
             expected_version: Optional[str] = None,
             from_step: Optional[str] = None,
             preflight_only: bool = False,
             switch_version_key: Optional[str] = None) -> dict:
    """from_step: 从第几步开始（1 开始的序号，或表单 ID），之前的步骤整步跳过。

    跳过的步骤连它的计算宏和运行包一起跳过（这些属于该步骤的收尾动作）。
    """
    flow = PbfFlow(engine, flow_cfg, form_cfgs, data_dir)
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
        if handle is not None and not handle.finished:
            handle.wait(float(login_cfg.get("timeout", 60)) + 30)
        if handle is not None and handle.finished and handle.outcome:
            log.info("启动阶段后台登录已成功，跳过重复登录")
            report["login"] = True
        else:
            if handle is not None and handle.finished and not handle.outcome \
                    and handle.saw_box:
                # 见过登录框但没登录成功：密码错误或登录框卡住，
                # 继续跑只会把清晰的报错变成 COM 挂死
                raise ExcelError(
                    f"AO 自动登录失败（{handle.error or '超时或密码错误'}），"
                    "登录框可能仍在——请手动登录后重试")
            from .ao_login import AoLogin
            ok = AoLogin(
                username=login_cfg["username"],
                password=login_cfg.get("password", ""),
                client=str(login_cfg.get("client", "333")),
                lang=str(login_cfg.get("lang", "zh")),
            ).ensure(float(login_cfg.get("timeout", 60)),
                     # 刚由本工具启动 Excel 时，登录框还没弹出来，必须等它出现
                     expect_login=bool(getattr(engine, "launched_now", False)))
            if not ok:
                raise ExcelError("AO 自动登录失败，请手动登录后重试")
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
                del report["steps"][steps_len:]      # 丢掉这次重试前的半成品记录
                del report["packages"][pkgs_len:]
                if attempt >= attempts or not is_disconnect(exc):
                    raise
                log.warning(f"{label} 因 Excel 断开失败：{exc}")
                if not flow.recover():
                    raise
                log.info(f"Excel 已恢复，重做：{label}")

    # 5) 切版本（用户路径第 ⑤ 步）：先切到选定版本，再「刷新状态」。
    switched_display: Optional[str] = None
    if switch_version_key:
        dims_file = app_root() / "config" / "dimensions.yaml"
        dims = yaml.safe_load(dims_file.read_text(encoding="utf-8")) or {}
        dims_nav = dims.get("navigator") or {}
        switched_display = (dims.get("versions") or {}).get(switch_version_key)
        run_with_recovery(
            lambda: Navigator(engine, dims_nav).switch_version(
                switch_version_key, switched_display),
            f"切换版本 {switch_version_key}")
        report["switched_version"] = switch_version_key

    expected = expected_version or switched_display
    report["version"] = run_with_recovery(
        lambda: flow.confirm_version(expected), "版本确认")
    flow.check_version_window(report["version"])
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
            for row in after_close.get("packages", []):
                r = flow.run_package(after_close["nav_cell"], row)
                report["packages"].append(r)
                if not r["ok"]:
                    raise ExcelError(f"运行包未成功: {r}")

    if only_verify:
        report["indicators"] = run_with_recovery(flow.read_indicators, "损益核对")
        return report

    for i, step in enumerate(steps):
        form_id = step["form"]
        if i < start_idx:
            log.info(f"跳过步骤 {i + 1}（{form_id}）")
            continue
        run_with_recovery(lambda s=step: run_step(s), f"步骤 {i + 1}（{form_id}）")

    report["indicators"] = run_with_recovery(flow.read_indicators, "损益核对")
    return report
