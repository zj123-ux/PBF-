"""PBF 滚动预测自动编制工具 — 本地 Web 界面。

Flask 只做「壳」：把网页参数落到磁盘（credentials.yaml / data_YYYY.MM），再用
subprocess 调现有 cli.py 驱动 Excel + AO，流式回传日志。单机单人使用，只绑 127.0.0.1。

启动:  python -m webapp.app    （或 run_web.bat）
端口:  默认 8765，可用环境变量 PBF_WEB_PORT 覆盖。
"""
from __future__ import annotations

import json
import os
import re
import subprocess
import sys
import threading
import time
import uuid
from datetime import datetime
from pathlib import Path
from urllib.parse import urlparse

# 确保项目根在 sys.path（兼容 `python webapp/app.py` 与 `python -m webapp.app`）
ROOT = Path(__file__).resolve().parent.parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

import yaml
from flask import Flask, jsonify, render_template, request

from paths import app_root
from data_dirs import latest_data_dir as find_latest_data_dir
from engine.business_units import UnitError, csv_dimension_warnings, default_unit, get_unit
from engine.periods import parse_periods
from webapp.runs import RunRegistry, should_hold_slot

APP_ROOT = app_root()
CONFIG_DIR = APP_ROOT / "config"
CREDENTIALS_FILE = CONFIG_DIR / "credentials.yaml"
DIMENSIONS_FILE = CONFIG_DIR / "dimensions.yaml"
BUSINESS_UNITS_FILE = CONFIG_DIR / "business_units.yaml"
FLOW_FILE = CONFIG_DIR / "flow.yaml"
LOG_DIR = APP_ROOT / "logs"
WEB_RUN_DIR = LOG_DIR / "web"

app = Flask(__name__)


@app.before_request
def _reject_cross_site_post():
    """只绑 127.0.0.1，但浏览器里别的站点仍可跨站 POST；按 Origin 拒绝。

    命令行/脚本调用通常不带 Origin，照常放行。
    """
    if request.method != "POST":
        return None
    origin = request.headers.get("Origin")
    if not origin:
        return None
    host = urlparse(origin).hostname or ""
    if host not in ("127.0.0.1", "localhost") and host != request.host.split(":")[0]:
        return jsonify({"ok": False, "error": "拒绝跨站 POST 请求"}), 403
    return None


# ---------- 配置读取 ----------

def _load_yaml(path: Path, default=None):
    if not path.exists():
        return default if default is not None else {}
    try:
        return yaml.safe_load(path.read_text(encoding="utf-8")) or {}
    except Exception:
        return default if default is not None else {}


def load_credentials() -> dict:
    return _load_yaml(CREDENTIALS_FILE, {})


def save_credentials(username: str, password: str, client: str, lang: str) -> None:
    CONFIG_DIR.mkdir(parents=True, exist_ok=True)
    header = ("# 本机凭据（不要提交到代码库、不要随项目一起发给别人）\n"
              "# 优先级：环境变量 PBF_USERNAME / PBF_PASSWORD / PBF_CLIENT / PBF_LANG > 本文件 > flow.yaml\n")
    body = yaml.safe_dump(
        {"username": username, "password": password, "client": client, "lang": lang},
        allow_unicode=True, sort_keys=False)
    CREDENTIALS_FILE.write_text(header + body, encoding="utf-8")


def load_versions() -> list[dict]:
    dims = _load_yaml(DIMENSIONS_FILE)
    versions = dims.get("versions") or {}
    return [{"key": k, "display": v} for k, v in versions.items()]


def load_business_units() -> dict:
    return _load_yaml(BUSINESS_UNITS_FILE)


def _resolve_unit(raw) -> tuple[str, str | None]:
    """校验前端传来的业务单元；未传则回退到第一个启用单元。

    返回 (unit_name, error)；error 非空表示请求非法。
    """
    units_cfg = load_business_units()
    if not units_cfg.get("units"):
        return "", None          # 未配置业务单元：按 flow.yaml 默认导航行运行
    name = str(raw or "").strip()
    if not name:
        name = default_unit(units_cfg) or ""
        if not name:
            return "", "没有启用的业务单元，请检查 config/business_units.yaml"
        return name, None
    try:
        get_unit(units_cfg, name)
    except UnitError as exc:
        return "", str(exc)
    return name, None


def latest_data_dir() -> str:
    """最新上传数据目录名（按期间排序，排除 *_clear）；没有返回空串。"""
    latest = find_latest_data_dir(APP_ROOT)
    return latest.name if latest else ""


def suggested_period() -> str:
    """最新数据目录的下一月，供前端默认填入。"""
    latest = latest_data_dir()
    m = re.search(r"(\d{4})\.(\d{2})", latest)
    if not m:
        return ""
    y, mo = int(m.group(1)), int(m.group(2))
    if mo == 12:
        return f"{y + 1}.01"
    return f"{y}.{mo + 1:02d}"


def _data_dir_name(periods: list[str]) -> str:
    """导出/运行目录名：单期间 data_YYYY.MM，多期间 data_<首>_<末>。"""
    first, last = periods[0], periods[-1]
    return f"data_{first}" if first == last else f"data_{first}_{last}"


def _parse_period_spec(spec: str) -> list[str]:
    """解析期间描述（单值/列表/区间），失败抛 ValueError。"""
    return parse_periods(spec)


# ---------- 子进程运行管理 ----------

# 单机单 Excel 实例，同一时刻只允许一个流程在跑（begin 原子抢占，避免竞态）
REGISTRY = RunRegistry()
RUN_TIMEOUT_ENV = "PBF_RUN_TIMEOUT_SEC"
# 默认不限时：财务长流程（多期间 + 三个运行包 + 落库校验）可能超过 1 小时，
# 强杀会把 Excel 停在保存/核对中途。需要限时用 PBF_RUN_TIMEOUT_SEC 显式设置
RUN_TIMEOUT_DEFAULT_SEC = 0


def _child_env() -> dict:
    env = os.environ.copy()
    # 让 cli.py 的 print 立即落盘，便于网页实时看到日志
    env["PYTHONUNBUFFERED"] = "1"
    # 子进程 stdout 是管道而不是控制台；固定 UTF-8，避免中文日志在父进程里乱码
    env["PYTHONIOENCODING"] = "utf-8"
    return env


def _run_timeout_sec() -> float | None:
    """单次子进程运行超时（秒）；PBF_RUN_TIMEOUT_SEC 未设置或<=0 表示不限时。"""
    try:
        value = float(os.environ.get(RUN_TIMEOUT_ENV, RUN_TIMEOUT_DEFAULT_SEC))
    except ValueError:
        value = float(RUN_TIMEOUT_DEFAULT_SEC)
    return value if value > 0 else None


def _kill_tree(pid: int) -> bool:
    """终止子进程及其后代（Windows taskkill /T），返回是否调用成功。"""
    try:
        r = subprocess.run(["taskkill", "/PID", str(pid), "/T", "/F"],
                           capture_output=True, text=True, timeout=30)
        return r.returncode == 0
    except Exception:
        return False


def _excel_running() -> bool:
    """Excel 是否仍在运行；检测失败时按"在运行"处理（宁可多占位，不可抢跑）。"""
    try:
        from engine.bootstrap import excel_pids
        return bool(excel_pids())
    except Exception:
        return True


def _try_release_stuck_run() -> None:
    """占位运行卡住时（kill 失败/Excel 未退出），条件解除后自动释放槽位。"""
    cur = REGISTRY.current
    if cur is None:
        return
    state = REGISTRY.get(cur)
    if not (state and state.get("stuck") and state.get("done")):
        return
    proc = state.get("proc")
    if proc is not None and proc.poll() is None:
        return
    if _excel_running():
        return
    state["stuck"] = False
    REGISTRY.end(cur)


def _busy_reason() -> str:
    """begin 失败时给用户的提示：区分"正常运行中"与"卡住的占位运行"。"""
    cur = REGISTRY.current
    state = REGISTRY.get(cur) if cur else None
    if state and state.get("stuck"):
        return ("上一次运行被取消/超时后未能确认 Excel 退出，运行槽位仍被占用；"
                "请手动关闭残留的 EXCEL.EXE 后重试（关闭后本提示自动解除）")
    return "已有一个流程在运行，请等待其结束"


def _run_step(cmd: list[str], lf, state: dict) -> int:
    """启动并等待一个 cli.py 子进程；支持取消与超时。"""
    if state.get("cancelled"):
        return -1
    try:
        proc = subprocess.Popen(
            cmd, cwd=str(APP_ROOT), env=_child_env(),
            stdout=lf, stderr=subprocess.STDOUT)
    except Exception as exc:
        lf.write(f"[Web] 无法启动子进程：{exc}\n")
        lf.flush()
        return -1
    state["proc"] = proc
    if state.get("cancelled"):
        _kill_tree(proc.pid)
    timeout = _run_timeout_sec()
    try:
        rc = proc.wait(timeout=timeout)
        if state.get("cancelled"):
            lf.write(f"[Web] 已按取消请求终止（退出码 {rc}）。\n")
            lf.flush()
        return rc
    except subprocess.TimeoutExpired:
        state["timed_out"] = True
        lf.write(f"[Web] 运行超过 {timeout:g}s 未结束，已终止子进程树。\n")
        lf.flush()
        _kill_tree(proc.pid)
        try:
            return proc.wait(timeout=20)
        except subprocess.TimeoutExpired:
            # taskkill 后子进程仍在：绝不放行新流程，槽位保持占用直到确认退出
            lf.write("[Web] 终止失败，子进程仍在运行；请手动结束该进程后重试。\n")
            lf.flush()
            return -1


def _spawn(params: dict) -> str | None:
    """启动一次后台运行，返回 run_id；已有流程在跑时返回 None。"""
    run_id = datetime.now().strftime("%Y%m%d_%H%M%S") + "_" + uuid.uuid4().hex[:6]
    WEB_RUN_DIR.mkdir(parents=True, exist_ok=True)
    log_path = WEB_RUN_DIR / f"{run_id}.log"
    report_path = WEB_RUN_DIR / f"{run_id}.json"

    period = params["period"]
    mode = params.get("mode", "fill")          # "fill" 填数 / "clear" 批量清除
    data_dir = params.get("data_dir")           # Path，fill 模式必填
    unit = params.get("unit") or ""             # 业务单元（品牌×品类）
    version_key = params.get("version_key") or ""
    version_display = params.get("version_display") or ""
    switch_version = bool(params.get("switch_version"))

    state = {"run_id": run_id, "done": False, "exit_code": None, "cancelled": False,
             "timed_out": False, "proc": None, "log_path": log_path,
             "report_path": report_path, "started": time.time()}
    _try_release_stuck_run()
    if not REGISTRY.begin(run_id, state):
        return None

    def worker() -> None:
        cli = str(APP_ROOT / "cli.py")
        try:
            with open(log_path, "w", encoding="utf-8") as lf:
                lf.write(f"[Web] 运行参数: 期间={period} 模式={mode}"
                         f" 业务单元={unit or '(默认)'}"
                         f" 版本={version_display or '(未指定)'} 切换版本={switch_version}\n")
                lf.flush()
                if switch_version and version_key:
                    lf.write(f"[Web] 切换版本 → {version_key}\n")
                    lf.flush()
                    rc = _run_step([sys.executable, cli, "switch-version", version_key], lf, state)
                    if rc != 0:
                        lf.write(f"[Web] 切换版本失败（退出码 {rc}），流程中止。\n")
                        lf.flush()
                        state["exit_code"] = rc
                        return
                if state.get("cancelled"):
                    lf.write("[Web] 已取消，流程未继续执行。\n")
                    lf.flush()
                    return
                if mode == "clear":
                    lf.write("[Web] 开始执行批量清除（生成全 0 并运行）…\n")
                    lf.flush()
                    cmd = [sys.executable, cli, "clear", "--periods", period,
                           "--run", "--report-file", str(report_path)]
                    if unit:
                        cmd += ["--unit", unit]
                else:
                    lf.write("[Web] 开始执行全流程（run-flow）…\n")
                    lf.flush()
                    cmd = [sys.executable, cli, "run-flow", "--data-dir", str(data_dir)]
                    if unit:
                        cmd += ["--unit", unit]
                    if version_display:
                        cmd += ["--version", version_display]
                    cmd += ["--report-file", str(report_path)]
                state["exit_code"] = _run_step(cmd, lf, state)
        except Exception as exc:
            if state["exit_code"] is None:
                state["exit_code"] = -1
            try:
                with open(log_path, "a", encoding="utf-8") as lf:
                    lf.write(f"[Web] 运行线程异常：{exc}\n")
            except OSError:
                pass
        finally:
            proc = state.get("proc")
            state["proc"] = None
            state["done"] = True
            # 取消/超时后只有确认子进程死亡且 Excel 已退出才释放槽位；
            # 否则两个流程会操作同一个 Excel。卡住的占位在条件解除后由
            # _try_release_stuck_run() 自动放行。
            proc_alive = proc is not None and proc.poll() is None
            if should_hold_slot(state, proc_alive, _excel_running()):
                state["stuck"] = True
                try:
                    with open(log_path, "a", encoding="utf-8") as lf:
                        lf.write("[Web] 未能确认子进程/Excel 完全退出，运行槽位继续保持"
                                 "占用；请手动关闭残留的 EXCEL.EXE，关闭后即可重新运行。\n")
                except OSError:
                    pass
            else:
                REGISTRY.end(run_id)

    threading.Thread(target=worker, daemon=True).start()
    return run_id


# ---------- 环境自检 ----------

def _proc_running(*names: str) -> bool:
    try:
        out = subprocess.run(["tasklist", "/FO", "CSV", "/NH"],
                             capture_output=True, text=True, timeout=15).stdout
    except Exception:
        return False
    upper = out.upper()
    return any(n.upper() in upper for n in names)


def check_env() -> dict:
    """被动自检，不启动 Excel、不改动任何状态。"""
    checks = []
    wps = _proc_running("et.exe", "wps.exe")
    checks.append({"ok": True, "label": "WPS 状态",
                   "detail": "WPS 表格在运行（已自动绕开，不影响 Excel 操作）" if wps
                             else "WPS 未运行"})

    excel_pids = []
    try:
        from engine.bootstrap import excel_pids
        excel_pids = excel_pids()
    except Exception:
        pass
    excel_on = bool(excel_pids)
    checks.append({"ok": excel_on, "label": "Excel 已运行",
                   "detail": f"Excel 进程数 {len(excel_pids)}" if excel_on
                             else "Excel 未运行（点运行时会尝试自动打开导航表）"})

    nav_open = False
    nav_detail = "未检测"
    if excel_on:
        try:
            import win32com.client
            xl = win32com.client.GetActiveObject("Excel.Application")
            if float(str(xl.Version)) < 14:
                nav_detail = "检测到 WPS 抢占 COM 接口（已自动排除）"
            else:
                flow = _load_yaml(FLOW_FILE)
                pattern = (flow.get("navigator") or {}).get(
                    "workbook_pattern", "PBF Pro - 全面预算系统流程导航*")
                prefix = pattern.rstrip("*")
                names = [w.Name for w in xl.Workbooks]
                if any(n.startswith(prefix) for n in names):
                    nav_open = True
                    nav_detail = "导航表已打开"
                else:
                    nav_detail = "真实 Excel 已运行，但导航表未打开（运行时会自动打开）"
        except Exception as exc:
            nav_detail = f"导航表未检测到（{str(exc)[:80]}）"
    checks.append({"ok": nav_open, "label": "AO 导航表",
                   "detail": nav_detail})

    return {"checks": checks, "excel_on": excel_on, "wps_on": wps}


# ---------- API ----------

@app.route("/")
def index():
    return render_template("index.html")


@app.route("/api/init")
def api_init():
    creds = load_credentials()
    units_cfg = load_business_units()
    units = []
    for name, u in (units_cfg.get("units") or {}).items():
        units.append({
            "name": name,
            "enabled": bool(u.get("启用", False)),
            "brand": u.get("品牌", ""),
            "category": u.get("品类", ""),
        })
    return jsonify({
        "versions": load_versions(),
        "brands": units_cfg.get("品牌", []),
        "categories": units_cfg.get("品类", []),
        "units": units,
        "default_unit": default_unit(units_cfg),
        "username": creds.get("username", ""),
        "password_set": bool(creds.get("password")),
        "client": str(creds.get("client", "333")),
        "lang": str(creds.get("lang", "zh")),
        "latest_data_dir": latest_data_dir(),
        "suggested_period": suggested_period(),
    })


@app.route("/api/credentials", methods=["POST"])
def api_credentials():
    data = request.get_json(force=True) or {}
    username = str(data.get("username", "")).strip()
    client = str(data.get("client", "333")).strip() or "333"
    lang = str(data.get("lang", "zh")).strip() or "zh"
    if not username:
        return jsonify({"ok": False, "error": "用户名不能为空"}), 400
    # 密码留空（或前端传 null）→ 保留已保存的密码，不覆盖成空
    raw_pw = data.get("password")
    if raw_pw in (None, ""):
        password = load_credentials().get("password", "")
    else:
        password = str(raw_pw)
    save_credentials(username, password, client, lang)
    return jsonify({"ok": True})


@app.route("/api/upload", methods=["POST"])
def api_upload():
    file = request.files.get("file")
    if file is None or file.filename == "":
        return jsonify({"ok": False, "error": "未选择文件"}), 400
    period = str(request.form.get("period", "")).strip()
    try:
        periods = _parse_period_spec(period)
    except ValueError as exc:
        return jsonify({"ok": False, "error": f"期间格式错误：{exc}"}), 400

    suffix = Path(file.filename).suffix.lower()
    if suffix not in (".xlsx", ".xlsm"):
        return jsonify({"ok": False, "error": f"仅支持 .xlsx / .xlsm，实际为 {suffix}"}), 400

    tmp = WEB_RUN_DIR / f"upload_{uuid.uuid4().hex[:8]}{suffix}"
    WEB_RUN_DIR.mkdir(parents=True, exist_ok=True)
    file.save(str(tmp))
    try:
        from engine.template_export import EmptyExportError, export_template
        out_dir = APP_ROOT / _data_dir_name(periods)
        # min_rows=1：0 行模板在替换目录之前就报错，不破坏上一份正确数据
        rows = export_template(tmp, periods, out_dir, min_rows=1)
    except EmptyExportError as exc:
        return jsonify({"ok": False, "error": str(exc)}), 400
    except Exception as exc:
        return jsonify({"ok": False, "error": f"导出失败：{exc}"}), 500
    finally:
        try:
            tmp.unlink()
        except Exception:
            pass

    total = sum(rows.values())
    return jsonify({"ok": True, "rows": rows, "total": total, "data_dir": out_dir.name,
                    "periods": periods,
                    "warnings": csv_dimension_warnings(out_dir)})


@app.route("/api/run", methods=["POST"])
def api_run():
    data = request.get_json(force=True) or {}
    period = str(data.get("period", "")).strip()
    try:
        periods = _parse_period_spec(period)
    except ValueError as exc:
        return jsonify({"ok": False, "error": f"期间格式错误：{exc}"}), 400
    unit, unit_error = _resolve_unit(data.get("unit"))
    if unit_error:
        return jsonify({"ok": False, "error": unit_error}), 400
    data_dir = APP_ROOT / _data_dir_name(periods)
    if not data_dir.exists():
        return jsonify({"ok": False, "error": f"数据目录不存在: {data_dir.name}，请先上传模板导出"}), 400
    run_id = _spawn({
        "period": period,
        "data_dir": data_dir,
        "unit": unit,
        "mode": "fill",
        "version_key": data.get("version_key") or "",
        "version_display": data.get("version_display") or "",
        "switch_version": bool(data.get("switch_version")),
    })
    if run_id is None:
        return jsonify({"ok": False, "error": _busy_reason()}), 409
    return jsonify({"ok": True, "run_id": run_id})


@app.route("/api/clear", methods=["POST"])
def api_clear():
    data = request.get_json(force=True) or {}
    period = str(data.get("period", "")).strip()
    try:
        periods = _parse_period_spec(period)
    except ValueError as exc:
        return jsonify({"ok": False, "error": f"期间格式错误：{exc}"}), 400
    unit, unit_error = _resolve_unit(data.get("unit"))
    if unit_error:
        return jsonify({"ok": False, "error": unit_error}), 400
    run_id = _spawn({
        "period": period,
        "unit": unit,
        "mode": "clear",
        "version_key": data.get("version_key") or "",
        "version_display": data.get("version_display") or "",
        "switch_version": bool(data.get("switch_version")),
    })
    if run_id is None:
        return jsonify({"ok": False, "error": _busy_reason()}), 409
    return jsonify({"ok": True, "run_id": run_id, "periods": periods})


@app.route("/api/run/<run_id>/log")
def api_run_log(run_id: str):
    state = REGISTRY.get(run_id)
    if state is None:
        return jsonify({"ok": False, "error": "未知 run_id"}), 404
    offset = int(request.args.get("offset", 0))
    try:
        content = state["log_path"].read_text(encoding="utf-8", errors="replace")
    except FileNotFoundError:
        content = ""
    new = content[offset:]
    # 退出码可能被 stderr 的报错码污染，这里统一取进程返回码
    exit_code = state["exit_code"] if state["done"] else None
    return jsonify({
        "ok": True, "done": state["done"], "exit_code": exit_code,
        "cancelled": bool(state.get("cancelled")),
        "timed_out": bool(state.get("timed_out")),
        "stuck": bool(state.get("stuck")),
        "lines": new, "next_offset": len(content),
    })


@app.route("/api/run/<run_id>/result")
def api_run_result(run_id: str):
    state = REGISTRY.get(run_id)
    if state is None:
        return jsonify({"ok": False, "error": "未知 run_id"}), 404
    if not state["done"]:
        return jsonify({"ok": False, "error": "流程尚未结束"}), 409
    report = {}
    if state["report_path"].exists():
        try:
            report = json.loads(state["report_path"].read_text(encoding="utf-8"))
        except Exception:
            report = {}
    return jsonify({"ok": True, "exit_code": state["exit_code"], "report": report,
                    "cancelled": bool(state.get("cancelled")),
                    "timed_out": bool(state.get("timed_out"))})


@app.route("/api/run/<run_id>/cancel", methods=["POST"])
def api_run_cancel(run_id: str):
    state = REGISTRY.get(run_id)
    if state is None:
        return jsonify({"ok": False, "error": "未知 run_id"}), 404
    if state["done"]:
        return jsonify({"ok": False, "error": "流程已结束，无法取消"}), 400
    state["cancelled"] = True
    proc = state.get("proc")
    if proc is not None and proc.poll() is None:
        _kill_tree(proc.pid)
        note = ("已终止子进程（Excel 由 COM 启动、不在子进程树内，taskkill 杀不到；"
                "若 Excel 未自动退出请手动关闭，运行槽位会保持占用直到 Excel 退出）")
    else:
        note = "已标记取消；当前步骤尚未启动子进程"
    return jsonify({"ok": True, "note": note})


@app.route("/api/env")
def api_env():
    return jsonify(check_env())


def cleanup_web_dir(max_age_days: int = 30) -> int:
    """删除 logs/web 下过期的日志/结果/上传临时文件，返回删除数量。"""
    if not WEB_RUN_DIR.exists():
        return 0
    cutoff = time.time() - max_age_days * 86400
    removed = 0
    for p in WEB_RUN_DIR.iterdir():
        try:
            if p.is_file() and p.stat().st_mtime < cutoff:
                p.unlink()
                removed += 1
        except OSError:
            pass
    return removed


def main() -> None:
    port = int(os.environ.get("PBF_WEB_PORT", "8765"))
    removed = cleanup_web_dir()
    if removed:
        print(f"已清理 {removed} 个过期网页运行日志")
    print(f"PBF Web 界面已启动: http://127.0.0.1:{port}")
    print("仅本机可访问（127.0.0.1）。按 Ctrl+C 退出。")
    app.run(host="127.0.0.1", port=port, debug=False, threaded=True)


if __name__ == "__main__":
    main()
