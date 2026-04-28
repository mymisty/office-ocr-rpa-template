from __future__ import annotations

from dataclasses import dataclass
from typing import Iterable

from .models import ClickTarget
from .safety import StopController


@dataclass
class ClickOptions:
    button: str = "left"
    clicks: int = 1
    offset_x: int = 0
    offset_y: int = 0
    move_duration: float = 0.15
    dry_run: bool = False


class Clicker:
    def __init__(self, dry_run: bool = False, fail_safe: bool = True, stop: StopController | None = None) -> None:
        import pyautogui

        self.pyautogui = pyautogui
        self.pyautogui.FAILSAFE = fail_safe
        self.dry_run = dry_run
        self.stop = stop

    def click(self, target: ClickTarget, options: ClickOptions | None = None) -> ClickTarget:
        options = options or ClickOptions(dry_run=self.dry_run)
        dry_run = self.dry_run or options.dry_run
        x = int(target.x + options.offset_x)
        y = int(target.y + options.offset_y)
        if self.stop:
            self.stop.check()
        if dry_run:
            return ClickTarget(x=x, y=y, box=target.box, description=f"dry-run: {target.description}")
        self.pyautogui.moveTo(x, y, duration=options.move_duration)
        if self.stop:
            self.stop.check()
        self.pyautogui.click(x=x, y=y, clicks=options.clicks, button=options.button)
        return ClickTarget(x=x, y=y, box=target.box, description=target.description)

    def click_xy(self, x: int, y: int, options: ClickOptions | None = None) -> ClickTarget:
        return self.click(ClickTarget(x=int(x), y=int(y), description="fixed coordinate"), options)

    def input_text(self, text: str, dry_run: bool | None = None) -> None:
        if self.stop:
            self.stop.check()
        if self.dry_run if dry_run is None else dry_run:
            return
        try:
            import pyperclip

            pyperclip.copy(text)
            self.pyautogui.hotkey("ctrl", "v")
        except Exception:
            self.pyautogui.write(text)

    def hotkey(self, keys: Iterable[str], dry_run: bool | None = None) -> None:
        if self.stop:
            self.stop.check()
        if self.dry_run if dry_run is None else dry_run:
            return
        self.pyautogui.hotkey(*list(keys))

    def scroll(self, amount: int, dry_run: bool | None = None) -> None:
        if self.stop:
            self.stop.check()
        if self.dry_run if dry_run is None else dry_run:
            return
        self.pyautogui.scroll(int(amount))
