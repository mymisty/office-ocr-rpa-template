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

    @property
    def center_x(self) -> int:
        return int((self.x1 + self.x2) / 2)

    @property
    def center_y(self) -> int:
        return int((self.y1 + self.y2) / 2)

    def to_dict(self) -> dict[str, Any]:
        data = asdict(self)
        data["center_x"] = self.center_x
        data["center_y"] = self.center_y
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
