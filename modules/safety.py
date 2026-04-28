from __future__ import annotations

import threading
import time
from dataclasses import dataclass


class StopRequested(RuntimeError):
    """Raised when the user asks the automation to stop."""


@dataclass
class StopController:
    hotkey: str = "esc"
    enabled: bool = True

    def __post_init__(self) -> None:
        self._stop_event = threading.Event()
        self._registered = False

    def start(self) -> None:
        if not self.enabled or self._registered:
            return
        try:
            import keyboard

            keyboard.add_hotkey(self.hotkey, self.request_stop)
            self._registered = True
        except Exception:
            self._registered = False

    def request_stop(self) -> None:
        self._stop_event.set()

    def check(self) -> None:
        if self._stop_event.is_set():
            raise StopRequested("User requested stop")

    def sleep(self, seconds: float) -> None:
        end = time.time() + seconds
        while time.time() < end:
            self.check()
            time.sleep(min(0.1, end - time.time()))

    def close(self) -> None:
        if not self._registered:
            return
        try:
            import keyboard

            keyboard.remove_hotkey(self.hotkey)
        except Exception:
            pass
        self._registered = False
