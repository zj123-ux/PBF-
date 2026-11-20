"""Excel COM engine — wraps pywin32 to talk to a running Excel + SAP AO.

All methods are thin wrappers that catch COM errors and raise structured
ExcelError exceptions so callers don't have to know pywin32 internals.
"""
from __future__ import annotations

import fnmatch
import logging
import os
import time
from typing import Any, Optional

try:
    import win32com.client  # type: ignore
    import win32com.server.exception  # type: ignore
except ImportError:  # pragma: no cover
    raise ImportError("pywin32 required: pip install pywin32")

log = logging.getLogger(__name__)

# 重要教训：WPS 不只是抢注 ProgID「Excel.Application」，它还会把真实 Excel 的 CLSID
# {00024500-0000-0000-C000-000000000046} 的 LocalServer32 改写成 et.exe。
# 因此用这个 CLSID 去 Dispatch/CoCreateInstance 会拿到 WPS（Version=12），而不是真 Excel。
# 正确做法：改用版本相关的 Excel.ChartApplication.<ver> CLSID
# {00024505-00<ver>-0000-C000-000000000046}（其 LocalServer32 仍指向真实 EXCEL.EXE），
# Dispatch 后校验 Version>=14 区分真假。见 _real_excel_clsids()。

# 缓存的真实 Excel 实例：让本进程内所有 ExcelEngine 共享同一个 Excel，
# 避免「ensure_excel_ready 建了一个专用实例、后续 ExcelEngine() 又连到另一个」。
_active_excel = None


def reset_active_excel() -> None:
    """失效缓存（在 kill_excel 结束 Excel 进程后调用）。"""
    global _active_excel
    _active_excel = None


class ExcelError(RuntimeError):
    """Raised for any COM / Excel-side failure."""


class RecoverableError(ExcelError):
    """Excel 没断线但已经卡死（如反复点"打开"不出表单）。

    和断线一样处理：结束本流程的 Excel 进程 → 重开导航表 → 重新登录 →
    从失败的步骤重做，而不是在卡住的窗口上无限重试。
    """


# Excel 进程退出/断电后，COM 调用会报这些错，需要重开 Excel 而不是简单重试
_DISCONNECT_MARKERS = (
    "已与其客户端断开连接",     # RPC_E_DISCONNECTED (-2147417848)
    "-2147417848",
    "远程过程调用失败",          # RPC_S_CALL_FAILED (-2147023170)
    "-2147023170",
    "RPC 服务器不可用",          # RPC_S_SERVER_UNAVAILABLE (-2147023174)
    "-2147023174",
    "服务器运行失败",            # CO_E_SERVER_EXEC_FAILURE (-2146959355)
    "-2146959355",
)


def is_disconnect(exc: BaseException) -> bool:
    """判断异常是否表示 Excel 连接已断开（可自动恢复）。"""
    text = str(exc)
    return any(marker in text for marker in _DISCONNECT_MARKERS)


def _macro_timeout_sec() -> float:
    """宏运行多久后开始告警（秒）。PBF_MACRO_TIMEOUT_SEC<=0 禁用。"""
    try:
        value = float(os.environ.get("PBF_MACRO_TIMEOUT_SEC", "900"))
    except ValueError:
        value = 900.0
    return value if value > 0 else 0.0


class ExcelEngine:
    """Connects to the active Excel process (must already be running with AO)."""

    def __init__(self, timeout_sec: float = 30.0) -> None:
        self._excel = self._connect(timeout_sec)
        # 记录本实例连接的是哪个 Excel 进程：断线恢复时只结束这个进程，
        # 不再无差别 taskkill 所有 EXCEL.EXE（避免杀掉用户其它未保存的工作簿）。
        self.pid = self._excel_pid(self._excel)
        # 最近一次宏执行期间被自动关闭的弹窗文本（供上层判断"没有要保存的数据"这类情况）
        self.last_dialogs: list[str] = []

    @staticmethod
    def _excel_pid(xl) -> Optional[int]:
        try:
            import win32process

            return int(win32process.GetWindowThreadProcessId(int(xl.Hwnd))[1])
        except Exception:
            return None

    # ---------- connection ----------

    @staticmethod
    def _version(xl) -> float:
        try:
            return float(str(xl.Version))
        except Exception:
            return 0.0

    @staticmethod
    def _is_real_excel(xl) -> bool:
        """真 Excel 的 Version>=14；WPS 伪装成 12.0。"""
        return ExcelEngine._version(xl) >= 14.0

    @staticmethod
    def _real_excel_path() -> Optional[str]:
        """定位真实 Excel 的 EXE 路径。

        WPS 会劫持 CLSID 的 LocalServer32，但劫持不了 App Paths 注册；优先从那里读。
        同时用它推断本机 Office 版本号（路径里的 Office16 → 16）。
        """
        candidates: list[str] = []
        try:
            import winreg
            with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE,
                                r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\excel.exe") as k:
                v, _ = winreg.QueryValueEx(k, "")
                if v:
                    candidates.append(v)
        except Exception:
            pass
        candidates += [
            r"C:\Program Files\Microsoft Office\Root\Office16\EXCEL.EXE",
            r"C:\Program Files (x86)\Microsoft Office\Root\Office16\EXCEL.EXE",
        ]
        for p in candidates:
            if p and os.path.exists(p):
                return p
        return None

    @staticmethod
    def _real_excel_clsids() -> list[str]:
        """返回 WPS 通常不劫持的、版本相关的 Excel CLSID 列表（本机版本优先）。

        背景：WPS 会把版本无关的 {00024500-0000-0000-C000-000000000046} 的 LocalServer32
        改成 et.exe，导致 Dispatch 那个 CLSID 拿到 WPS。但版本相关的
        Excel.ChartApplication.<ver>（{00024505-00<ver>-0000-C000-000000000046}）
        的 LocalServer32 仍指向真实 EXCEL.EXE，用它 Dispatch 拿到的就是真 Excel。
        """
        import re
        ids: list[str] = []
        path = ExcelEngine._real_excel_path() or ""
        m = re.search(r"Office(\d+)", path, re.IGNORECASE)
        if m:
            try:
                ids.append(f"{{00024505-{int(m.group(1)):04d}-0000-C000-000000000046}}")
            except Exception:
                pass
        for v in (16, 15, 14):
            c = f"{{00024505-{v:04d}-0000-C000-000000000046}}"
            if c not in ids:
                ids.append(c)
        return ids

    @staticmethod
    def _dispatch_real_excel():
        """Dispatch 版本相关 CLSID，拿到真实 Excel（Version>=14）。

        Dispatch 会复用已在运行的 Excel（单实例），没运行时会自动启动它。
        拿到 WPS（Version=12）或失败则返回 None，绝不调用 Quit()（那会关掉用户的 WPS）。
        """
        for clsid in ExcelEngine._real_excel_clsids():
            try:
                xl = win32com.client.Dispatch(clsid)
            except Exception:
                continue
            if ExcelEngine._is_real_excel(xl):
                return xl
        return None

    @staticmethod
    def _connect(timeout: float):
        global _active_excel
        # 已有缓存的真实 Excel → 直接复用（保证全流程操作同一个实例）
        if _active_excel is not None:
            try:
                if ExcelEngine._is_real_excel(_active_excel):
                    return _active_excel
            except Exception:
                pass
            _active_excel = None

        # 连接真实 Excel（复用运行中的，或启动一个专用实例）；短暂重试应对启动慢。
        deadline = time.time() + min(timeout, 5.0)
        while time.time() < deadline:
            xl = ExcelEngine._dispatch_real_excel()
            if xl is not None:
                try:
                    xl.Visible = True
                except Exception:
                    pass
                _active_excel = xl
                return xl
            time.sleep(0.5)
        raise ExcelError("无法连接真实 Excel：相关 CLSID 均被 WPS 抢占。"
                         "请关闭 WPS 表格后重试，或修复 Office 的 COM 注册")

    # ---------- workbook lookup ----------

    def find_workbook(self, pattern: str) -> Any:
        """Find open workbook whose Name (with timestamp suffix) matches `pattern`.

        Example pattern: "YCCBW028_其他成本编制表_CBG*"
        Matches files like "YCCBW028_其他成本编制表_CBG (15-15-23).xlsb".
        """
        wbs = list(self._excel.Workbooks)
        for wb in wbs:
            if fnmatch.fnmatchcase(wb.Name, pattern):
                return wb
        names = [w.Name for w in wbs]
        raise ExcelError(f"No workbook matching '{pattern}'. Open: {names}")

    def get_workbook(self, pattern: str, wait_sec: float = 20.0) -> Any:
        """轮询查找工作簿。AO 忙时会抛 RPC_E_CALL_REJECTED，这里统一重试。"""
        deadline = time.time() + wait_sec
        while time.time() < deadline:
            try:
                return self.find_workbook(pattern)
            except Exception:
                time.sleep(0.5)
        try:
            return self.find_workbook(pattern)
        except Exception as exc:
            raise ExcelError(f"未找到工作簿 '{pattern}'（等待 {wait_sec}s）: {exc}") from exc

    def open_sheet(self, wb: Any, sheet_name: str) -> Any:
        for sh in wb.Worksheets:
            if sh.Name == sheet_name:
                return sh
        names = [s.Name for s in wb.Worksheets]
        raise ExcelError(f"No sheet '{sheet_name}' in {wb.Name}. Sheeets: {names}")

    # ---------- cell / range ----------

    # ---------- 事件开关 ----------

    def events_enabled(self) -> Optional[bool]:
        """读取 Application.EnableEvents；读不到返回 None。"""
        try:
            return bool(self._excel.EnableEvents)
        except Exception:
            return None

    def ensure_events(self) -> bool:
        """确保 EnableEvents=True。返回 True 表示之前是关闭的（被强制打开）。"""
        try:
            if not bool(self._excel.EnableEvents):
                self._excel.EnableEvents = True
                return True
        except Exception:
            pass
        return False

    def read(self, sheet: Any, cell: str) -> Any:
        try:
            v = sheet.Range(cell).Value2
            # Excel COM often returns a tuple-of-tuples for a single cell too — unwrap
            if isinstance(v, tuple) and len(v) == 1 and isinstance(v[0], tuple):
                v = v[0][0]
            return v
        except Exception as exc:
            raise ExcelError(f"read {sheet.Name}!{cell}: {exc}") from exc

    def write(self, sheet: Any, cell: str, value: Any) -> None:
        try:
            sheet.Range(cell).Value2 = value
        except Exception as exc:
            raise ExcelError(f"write {sheet.Name}!{cell}={value!r}: {exc}") from exc

    def write_batch(self, sheet: Any, updates: dict[str, Any]) -> None:
        """Write many cells in one shot — one COM call per cell, but batched in Python."""
        errors = []
        for cell, value in updates.items():
            try:
                sheet.Range(cell).Value2 = value
            except Exception as exc:
                errors.append(f"{cell}={value!r}: {exc}")
        if errors:
            raise ExcelError("Batch write failed:\n" + "\n".join(errors))

    def copy_paste(self, sheet: Any, source_range: str, target_range: str) -> None:
        """Source → target via Excel Clipboard (matches how humans paste)."""
        try:
            sheet.Range(source_range).Copy()
            sheet.Range(target_range).PasteSpecial(Paste=-4122)  # xlPasteAll
            self._excel.CutCopyMode = False
        except Exception as exc:
            raise ExcelError(f"copy-paste {source_range} → {target_range}: {exc}") from exc

    # ---------- selection ----------

    def select(self, sheet: Any, cell: str) -> None:
        try:
            # ActiveSheet 是只读属性，只能通过 Activate 切换
            sheet.Activate()
            sheet.Range(cell).Select()
        except Exception as exc:
            raise ExcelError(f"select {sheet.Name}!{cell}: {exc}") from exc

    # ---------- macros / AO automation ----------

    def run_macro(self, macro_name: str, wb: Optional[Any] = None, *args) -> Any:
        """Run a VBA macro. If wb given, qualify as wb.Macro; else global Run.

        AO PBF templates expose macros like Plan_Save, Refresh_Sheet,
        Trigger_PS, Switch_mode, etc. — names vary per template.
        执行期间会起一个看门狗线程，自动关掉 PBF 的提示弹窗（否则宏调用会一直阻塞）。
        """
        full = macro_name
        try:
            if wb is not None:
                # AO 的模板宏（Refresh_Sheet / Plan_Save 等）要求工作簿是活动窗口，
                # 否则可能静默不生效（实测：不激活时刷新不会重算期间列）。
                try:
                    self._retry(lambda: (wb.Activate(), None), attempts=3, delay=1.0)
                    time.sleep(0.5)
                except Exception:
                    pass
                # VBAProject references wb code-name or ThisWorkbook
                # Use Application.Run with workbook-qualified name
                name = self._retry(lambda: wb.Name)
                full = f"'{name}'!{macro_name}"
            return self._run_with_dialog_guard(
                lambda: self._retry(lambda: self._excel.Run(full, *args)), label=full)
        except Exception as exc:
            raise ExcelError(f"macro {full}: {exc}") from exc

    # ---------- 弹窗看门狗 ----------

    @staticmethod
    def _iter_dialogs(pid: int) -> list[int]:
        """列出某进程的模态对话框窗口（class #32770）。"""
        import win32gui
        import win32process

        found: list[int] = []

        def cb(hwnd, _):
            _, p = win32process.GetWindowThreadProcessId(hwnd)
            if p == pid and win32gui.IsWindowVisible(hwnd) and win32gui.GetClassName(hwnd) == "#32770":
                found.append(hwnd)
            return True

        win32gui.EnumWindows(cb, None)
        return found

    @staticmethod
    def _dialog_text(hwnd: int) -> str:
        import win32gui

        parts = [win32gui.GetWindowText(hwnd)]

        def cb(child, _):
            t = win32gui.GetWindowText(child)
            if t:
                parts.append(t)
            return True

        win32gui.EnumChildWindows(hwnd, cb, None)
        return " | ".join(p for p in parts if p)

    @staticmethod
    def _press_default_button(hwnd: int) -> bool:
        import win32con
        import win32gui

        buttons: list[int] = []

        def cb(child, _):
            if win32gui.GetClassName(child) == "Button" and win32gui.IsWindowVisible(child):
                buttons.append(child)
            return True

        win32gui.EnumChildWindows(hwnd, cb, None)
        for b in buttons:
            text = win32gui.GetWindowText(b).replace("&", "").strip()
            if text in ("确定", "OK", "是"):
                win32gui.SendMessage(b, win32con.BM_CLICK, 0, 0)
                return True
        if buttons:
            win32gui.SendMessage(buttons[0], win32con.BM_CLICK, 0, 0)
            return True
        return False

    def start_dialog_guard(self) -> None:
        """整个流程期间在后台关闭已知的 PBF 提示弹窗。

        导航表点"打开"、表单刷新时也可能弹模态框；弹窗不关，外部 COM 调用会一直
        阻塞（实测卡了 5 分钟以上），日志里只看到"未打开工作簿，重试"。
        """
        if getattr(self, "_guard_stop", None) is not None:
            return
        import threading

        try:
            import win32process

            pid = int(win32process.GetWindowThreadProcessId(int(self._excel.Hwnd))[1])
        except Exception as exc:
            log.debug(f"弹窗看门狗未启动: {exc}")
            return
        stop = threading.Event()
        threading.Thread(target=self._dismiss_dialogs_loop,
                         args=(pid, stop, []), daemon=True).start()
        self._guard_stop = stop

    def stop_dialog_guard(self) -> None:
        stop = getattr(self, "_guard_stop", None)
        if stop is not None:
            stop.set()
            self._guard_stop = None

    def _dismiss_dialogs_loop(self, pid: int, stop, seen: list[str]) -> None:
        """后台线程：关闭 PBF 的提示弹窗，并记录所有模态弹窗。

        只有已知的提示类弹窗会被自动点掉；不认识的弹窗一律不碰，但会记录
        文本——宏卡住时日志里至少能看到是哪个弹窗挡住了。
        """
        known = ("没有要保存", "请检查", "PBF")
        while not stop.is_set():
            try:
                for hwnd in self._iter_dialogs(pid):
                    text = self._dialog_text(hwnd)
                    recognized = any(k in text for k in known)
                    if text not in seen:
                        seen.append(text)
                        if recognized:
                            log.warning(f"检测到弹窗并自动关闭: {text}")
                        else:
                            log.warning(
                                f"检测到未识别的模态弹窗（未自动关闭，可能阻塞流程）: {text}")
                    if recognized:
                        self._press_default_button(hwnd)
            except Exception:
                pass
            time.sleep(0.4)

    def _run_with_dialog_guard(self, fn, label: str = ""):
        import threading

        self.last_dialogs = []
        stop = threading.Event()
        try:
            import win32process

            pid = win32process.GetWindowThreadProcessId(int(self._excel.Hwnd))[1]
            threading.Thread(target=self._dismiss_dialogs_loop,
                             args=(pid, stop, self.last_dialogs), daemon=True).start()
        except Exception as exc:
            log.debug(f"弹窗看门狗启动失败（不影响主流程）: {exc}")

        started = time.time()
        limit = _macro_timeout_sec()

        def watchdog() -> None:
            """COM 调用无法从外部安全打断，这里只负责在超时后持续告警。"""
            warned_at = 0.0
            while not stop.wait(5.0):
                elapsed = time.time() - started
                if limit and elapsed >= limit and elapsed - warned_at >= 60:
                    warned_at = elapsed
                    dialogs = "；".join(self.last_dialogs[-3:]) or "无"
                    log.error(
                        f"宏 {label or '(未命名)'} 已运行 {elapsed:.0f}s（阈值 {limit:.0f}s），"
                        f"可能被弹窗或 BPC 请求阻塞。近期弹窗：{dialogs}。"
                        "请在 Excel 中处理后重试，或在 Web 界面点「取消」终止流程。")

        threading.Thread(target=watchdog, daemon=True).start()
        try:
            return fn()
        finally:
            stop.set()

    def save(self, wb: Any, macro_hint: Optional[str] = None) -> Any:
        """Save via Plan_Save / custom macro — fall back to SaveData command."""
        # PBF templates usually have their own save macro
        if macro_hint:
            try:
                return self.run_macro(macro_hint, wb)
            except ExcelError:
                pass
        # Generic AO API
        try:
            return self._excel.Run("SAPExecuteCommand", "SaveData")
        except Exception as exc:
            raise ExcelError(f"save failed: {exc}") from exc

    def refresh(self, wb: Any, macro_hint: Optional[str] = None) -> Any:
        if macro_hint:
            try:
                return self.run_macro(macro_hint, wb)
            except ExcelError:
                pass
        try:
            return self._excel.Run("SAPExecuteCommand", "RefreshAll")
        except Exception as exc:
            raise ExcelError(f"refresh failed: {exc}") from exc

    # ---------- window ----------

    @staticmethod
    def _retry(fn, attempts: int = 6, delay: float = 2.0):
        """AO 计算期间会抛 RPC_E_CALL_REJECTED（被呼叫方拒绝接收呼叫），重试即可。"""
        last = None
        for _ in range(attempts):
            try:
                return fn()
            except Exception as exc:
                last = exc
                msg = str(exc)
                if not ("拒绝接收呼叫" in msg or "-2147418111" in msg or "Call was rejected" in msg):
                    raise
                time.sleep(delay)
        raise last

    def click_cell(self, sheet: Any, cell: str, x_frac: float = 0.5) -> bool:
        """用真实鼠标事件点击某个单元格（AO 的 [+]/[-] 折叠图标只能这样触发）。

        坐标通过 ActiveWindow.RangeFromPoint 反查屏幕位置，避免手工换算缩放/冻结窗格。
        x_frac 是点击位置在单元格宽度上的比例（表头文字居中，标记不一定在左边缘）。
        返回是否确认点到了目标单元格。
        """
        import ctypes

        user32 = ctypes.windll.user32
        xl = self._excel
        hwnd = int(xl.Hwnd)
        try:
            user32.ShowWindow(hwnd, 3)            # SW_MAXIMIZE
            # 仅 SetForegroundWindow 常被前台锁拦截，先模拟一次 Alt 按键解锁
            user32.keybd_event(0x12, 0, 0, 0)
            user32.keybd_event(0x12, 0, 2, 0)
            user32.SetForegroundWindow(hwnd)
        except Exception:
            pass
        time.sleep(1.0)

        try:
            self._retry(lambda: (sheet.Activate(), sheet.Range(cell).Select()))
        except Exception as exc:
            raise ExcelError(f"定位单元格 {cell} 失败: {exc}") from exc
        time.sleep(0.4)

        win = xl.ActiveWindow
        target = str(self._retry(lambda: sheet.Range(cell).Address))
        target_row = str(sheet.Range(cell).Row)

        def hit(x: int, y: int) -> str:
            try:
                r = win.RangeFromPoint(x, y)
                return str(r.Address) if r is not None else ""
            except Exception:
                return ""

        def cell_row(addr: str) -> str:
            # "$AT$52" → "52"
            return addr.rsplit("$", 1)[-1] if addr else ""

        # 先只按"行号"定位纵向位置（单元格可能远在右侧，不能要求整格命中）
        row_ys: list[int] = []
        for sample_x in (900, 700, 1100, 500, 1300):
            row_ys = [y for y in range(60, 1100, 3) if cell_row(hit(sample_x, y)) == target_row]
            if row_ys:
                break
        if not row_ys:
            raise ExcelError(f"屏幕上找不到单元格 {cell}（请确认 Excel 窗口未最小化）")
        row_y = row_ys[len(row_ys) // 2]      # 取竖直中点，避免点在行边界上

        xs = [x for x in range(0, 2000, 4) if hit(x, row_y) == target]
        if not xs:
            # 列在可视区之外：用 Select 滚动一次再找
            self._retry(lambda: (sheet.Range(cell).Select(), None))
            time.sleep(0.8)
            xs = [x for x in range(0, 2000, 4) if hit(x, row_y) == target]
        if not xs:
            raise ExcelError(f"屏幕上找不到单元格 {cell} 的水平范围")

        x = int(xs[0] + (xs[-1] - xs[0]) * x_frac)
        user32.SetCursorPos(x, row_y)
        time.sleep(0.25)
        user32.mouse_event(0x0002, 0, 0, 0, 0)    # LEFTDOWN
        time.sleep(0.08)
        user32.mouse_event(0x0004, 0, 0, 0, 0)    # LEFTUP
        return True

    def workbook_names(self) -> list[str]:
        return [wb.Name for wb in self._excel.Workbooks]

    def workbook_count(self) -> int:
        return self._excel.Workbooks.Count

    def close_workbook(self, wb: Any, save: bool = False) -> None:
        """关闭工作簿。默认丢弃本地改动（数据已通过 Plan_Save 落入 BPC）。

        AO 在关闭时会弹"是否保存工作簿"，必须临时关闭 DisplayAlerts 才能静默跳过。
        """
        name = getattr(wb, "Name", "?")
        old = None
        try:
            old = self._excel.DisplayAlerts
            self._excel.DisplayAlerts = False
            wb.Close(save)
        except Exception as exc:
            raise ExcelError(f"close {name}: {exc}") from exc
        finally:
            if old is not None:
                try:
                    self._excel.DisplayAlerts = old
                except Exception:
                    pass
