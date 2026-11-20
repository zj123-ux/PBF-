"""运行注册表：同一时刻只允许一个流程，并保留最近若干次运行状态。

独立于 Flask / engine，方便单元测试；线程安全。
"""
from __future__ import annotations

import threading

MAX_KEPT_RUNS = 20


class RunRegistry:
    """进程内运行状态表。

    begin() 原子地抢占「当前运行」槽位；end() 释放。只有持有槽位的运行才能
    被注册，从而避免两个请求同时通过「是否在运行」检查后再各自启动子进程。
    """

    def __init__(self, max_kept: int = MAX_KEPT_RUNS):
        self._runs: dict[str, dict] = {}
        self._lock = threading.RLock()
        self._current: str | None = None
        self._max_kept = max_kept

    def begin(self, run_id: str, state: dict) -> bool:
        """注册一次运行；已有运行在进行中时返回 False（不注册）。"""
        with self._lock:
            if self._current is not None:
                return False
            self._current = run_id
            self._runs[run_id] = state
            self._prune_locked()
            return True

    def end(self, run_id: str) -> None:
        with self._lock:
            if self._current == run_id:
                self._current = None

    def get(self, run_id: str) -> dict | None:
        with self._lock:
            return self._runs.get(run_id)

    @property
    def current(self) -> str | None:
        with self._lock:
            return self._current

    def _prune_locked(self) -> None:
        """只淘汰已结束的最旧运行，保持内存与日志引用不无限增长。"""
        if len(self._runs) <= self._max_kept:
            return
        finished = [rid for rid, s in self._runs.items()
                    if s.get("done") and rid != self._current]
        for rid in finished[:len(self._runs) - self._max_kept]:
            self._runs.pop(rid, None)


def should_hold_slot(state: dict, proc_alive: bool, excel_running: bool) -> bool:
    """取消/超时后是否继续占用运行槽位。

    占位的目的：子进程没被确认终止，或 COM 启动的 Excel 还在跑（它不属于
    子进程树，taskkill /T 杀不到），此时放行新流程会出现两个流程操作同一
    个 Excel。子进程确认死亡、Excel 也退出后才放行。
    """
    if proc_alive:
        return True
    if state.get("timed_out") or state.get("cancelled"):
        return excel_running
    return False


def can_force_release(state: dict, proc_alive: bool) -> bool:
    """是否允许人工强制释放槽位。

    只放行"运行已结束、被标记 stuck、子进程确已退出"的状态；仍在运行的
    子进程一律拒绝，避免人工操作放行后两个流程同时操作 Excel。
    用户点击前需自行确认 Excel 中没有仍在执行的宏。
    """
    return bool(state.get("done") and state.get("stuck")) and not proc_alive
