"""
Background system monitor for Jarvis.

Polls system state every 60s and pushes Alert objects to a queue that the
main loop drains between wake-word listens. Each alert type has its own
throttle so the assistant cannot spam the user, and quiet hours suppress
non-critical alerts overnight.
"""

from __future__ import annotations

import datetime as dt
import logging
import queue
import threading
import time
from dataclasses import dataclass

try:
    import psutil
except ImportError:  # pragma: no cover - validated at install time
    psutil = None  # type: ignore


# -- thresholds --------------------------------------------------------------

BATTERY_LOW = 30        # %
BATTERY_CRITICAL = 15   # %
DISK_LOW_PCT = 10       # % free on system drive
SCREEN_TIME_INTERVAL = 90 * 60  # 90 minutes
POLL_SECONDS = 60

# throttles per alert type, in seconds
THROTTLE = {
    "battery_low": 30 * 60,
    "battery_critical": 15 * 60,
    "disk_low": 6 * 60 * 60,
    "screen_time": 90 * 60,
    "morning": 24 * 60 * 60,
    "lunch": 24 * 60 * 60,
    "late_night": 24 * 60 * 60,
}

# quiet hours: midnight → 7am, only "high" priority alerts get through
QUIET_START_HOUR = 0
QUIET_END_HOUR = 7


@dataclass
class Alert:
    type: str
    message: str
    priority: str = "normal"  # "low" | "normal" | "high"
    ts: float = 0.0


class Monitor:
    def __init__(self, logger: logging.Logger, alert_queue: "queue.Queue[Alert]") -> None:
        self.logger = logger
        self.queue = alert_queue
        self._thread: threading.Thread | None = None
        self._stop = threading.Event()
        self._last_alert: dict[str, float] = {}
        self._started_at = time.time()
        self._last_screen_alert = self._started_at
        self._greeted_today: dict[str, str] = {}  # alert type -> date string

    # -- lifecycle -----------------------------------------------------------

    def start(self) -> None:
        if psutil is None:
            self.logger.info("MONITOR_DISABLED | psutil not installed")
            return
        if self._thread and self._thread.is_alive():
            return
        self._stop.clear()
        self._thread = threading.Thread(target=self._loop, name="jarvis-monitor", daemon=True)
        self._thread.start()
        self.logger.info("MONITOR_START | poll=%ds", POLL_SECONDS)

    def stop(self) -> None:
        self._stop.set()
        if self._thread:
            self._thread.join(timeout=2)
        self.logger.info("MONITOR_STOP")

    # -- live snapshot (used by briefing builder) ----------------------------

    def snapshot(self) -> dict:
        if psutil is None:
            return {}
        snap: dict = {}
        try:
            b = psutil.sensors_battery()
            if b is not None:
                snap["battery_percent"] = int(b.percent)
                snap["battery_plugged"] = bool(b.power_plugged)
        except Exception:
            pass
        try:
            snap["cpu_percent"] = psutil.cpu_percent(interval=None)
        except Exception:
            pass
        try:
            snap["ram_percent"] = psutil.virtual_memory().percent
        except Exception:
            pass
        try:
            d = psutil.disk_usage("C:\\")
            snap["disk_free_gb"] = round(d.free / (1024 ** 3), 1)
            snap["disk_free_pct"] = round(100 - d.percent, 1)
        except Exception:
            pass
        return snap

    # -- main loop -----------------------------------------------------------

    def _loop(self) -> None:
        while not self._stop.is_set():
            try:
                self._check_battery()
                self._check_disk()
                self._check_screen_time()
                self._check_time_of_day()
            except Exception as e:
                self.logger.info("MONITOR_ERROR | %s", e)
            self._stop.wait(POLL_SECONDS)

    # -- checks --------------------------------------------------------------

    def _check_battery(self) -> None:
        try:
            b = psutil.sensors_battery()
        except Exception:
            return
        if b is None:
            return
        if b.power_plugged:
            return
        pct = int(b.percent)
        if pct <= BATTERY_CRITICAL:
            self._emit(Alert(
                type="battery_critical",
                message=f"Battery is at {pct}%, unplugged. Critical level.",
                priority="high",
            ))
        elif pct <= BATTERY_LOW:
            self._emit(Alert(
                type="battery_low",
                message=f"Battery is at {pct}%, unplugged.",
                priority="normal",
            ))

    def _check_disk(self) -> None:
        try:
            d = psutil.disk_usage("C:\\")
        except Exception:
            return
        free_pct = 100 - d.percent
        if free_pct < DISK_LOW_PCT:
            free_gb = round(d.free / (1024 ** 3), 1)
            self._emit(Alert(
                type="disk_low",
                message=f"System drive has only {free_gb} GB free ({free_pct:.0f}%).",
                priority="normal",
            ))

    def _check_screen_time(self) -> None:
        now = time.time()
        if now - self._last_screen_alert >= SCREEN_TIME_INTERVAL:
            self._last_screen_alert = now
            mins = int((now - self._started_at) / 60)
            self._emit(Alert(
                type="screen_time",
                message=f"You've been working for {mins} minutes. A short break is advisable.",
                priority="low",
            ))

    def _check_time_of_day(self) -> None:
        # Give the startup greeting time to finish before daily reminders kick in.
        if time.time() - self._started_at < 300:
            return

        now = dt.datetime.now()
        today = now.strftime("%Y-%m-%d")
        h = now.hour

        if 6 <= h < 11 and self._greeted_today.get("morning") != today:
            self._greeted_today["morning"] = today
            self._emit(Alert(
                type="morning",
                message=f"Good morning. It is {now.strftime('%I:%M %p')}.",
                priority="low",
            ))
        elif 12 <= h < 14 and self._greeted_today.get("lunch") != today:
            self._greeted_today["lunch"] = today
            self._emit(Alert(
                type="lunch",
                message="It is around lunchtime. A meal might be in order.",
                priority="low",
            ))
        elif h >= 23 and self._greeted_today.get("late_night") != today:
            self._greeted_today["late_night"] = today
            self._emit(Alert(
                type="late_night",
                message=f"It is {now.strftime('%I:%M %p')}. Perhaps consider winding down.",
                priority="low",
            ))

    # -- emit / throttle / quiet hours --------------------------------------

    def _emit(self, alert: Alert) -> None:
        now = time.time()
        # throttle
        last = self._last_alert.get(alert.type, 0)
        cooldown = THROTTLE.get(alert.type, 30 * 60)
        if now - last < cooldown:
            return
        # quiet hours — only "high" priority gets through
        h = dt.datetime.now().hour
        if QUIET_START_HOUR <= h < QUIET_END_HOUR and alert.priority != "high":
            self.logger.info("ALERT_SUPPRESSED | quiet_hours | %s", alert.type)
            return
        self._last_alert[alert.type] = now
        alert.ts = now
        self.queue.put(alert)
        self.logger.info("ALERT_EMIT | %s | %s", alert.type, alert.message)
