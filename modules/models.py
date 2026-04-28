from __future__ import annotations

from dataclasses import asdict, dataclass
from pathlib import Path
from typing import Any


@dataclass(slots=True)
class OCRBox:
    text: str
    x1: int
    y1: int
    x2: int
    y2: int
    confidence: float
    source_image: str | None = None
    screen_offset_x: int = 0
    screen_offset_y: int = 0

    @property
    def center_x(self) -> int:
        return int((self.x1 + self.x2) / 2)

    @property
    def center_y(self) -> int:
        return int((self.y1 + self.y2) / 2)

    @property
    def screen_x1(self) -> int:
        return self.x1 + self.screen_offset_x

    @property
    def screen_y1(self) -> int:
        return self.y1 + self.screen_offset_y

    @property
    def screen_x2(self) -> int:
        return self.x2 + self.screen_offset_x

    @property
    def screen_y2(self) -> int:
        return self.y2 + self.screen_offset_y

    @property
    def screen_center_x(self) -> int:
        return self.center_x + self.screen_offset_x

    @property
    def screen_center_y(self) -> int:
        return self.center_y + self.screen_offset_y

    def to_dict(self) -> dict[str, Any]:
        data = asdict(self)
        data["center_x"] = self.center_x
        data["center_y"] = self.center_y
        data["screen_x1"] = self.screen_x1
        data["screen_y1"] = self.screen_y1
        data["screen_x2"] = self.screen_x2
        data["screen_y2"] = self.screen_y2
        data["screen_center_x"] = self.screen_center_x
        data["screen_center_y"] = self.screen_center_y
        return data


@dataclass(slots=True)
class ClickTarget:
    x: int
    y: int
    box: OCRBox | None = None
    description: str = ""


@dataclass(slots=True)
class RuntimeContext:
    run_id: str
    root_dir: Path
    dry_run: bool
    current_item: dict[str, Any] | None = None
    last_screenshot: Path | None = None
    last_ocr: list[OCRBox] | None = None
