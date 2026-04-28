from __future__ import annotations

import ctypes
from datetime import datetime
from pathlib import Path
from typing import Any

from .models import OCRBox


def set_dpi_awareness() -> None:
    if not hasattr(ctypes, "windll"):
        return
    try:
        ctypes.windll.user32.SetProcessDPIAware()
    except Exception:
        pass


def normalize_region(region: dict[str, Any] | list[int] | tuple[int, ...] | None) -> dict[str, int] | None:
    if region is None:
        return None
    if isinstance(region, dict):
        if {"left", "top", "width", "height"} <= set(region):
            return {
                "left": int(region["left"]),
                "top": int(region["top"]),
                "width": int(region["width"]),
                "height": int(region["height"]),
            }
        if {"x", "y", "width", "height"} <= set(region):
            return {
                "left": int(region["x"]),
                "top": int(region["y"]),
                "width": int(region["width"]),
                "height": int(region["height"]),
            }
        if {"x1", "y1", "x2", "y2"} <= set(region):
            return {
                "left": int(region["x1"]),
                "top": int(region["y1"]),
                "width": int(region["x2"]) - int(region["x1"]),
                "height": int(region["y2"]) - int(region["y1"]),
            }
    if len(region) == 4:
        x1, y1, x2, y2 = [int(v) for v in region]
        return {"left": x1, "top": y1, "width": x2 - x1, "height": y2 - y1}
    raise ValueError(f"Unsupported screenshot region: {region}")


class ScreenshotManager:
    def __init__(self, root_dir: str | Path, run_id: str) -> None:
        self.root_dir = Path(root_dir)
        self.run_id = run_id
        set_dpi_awareness()

    def capture(
        self,
        label: str,
        subdir: str = "before",
        region: dict[str, Any] | list[int] | tuple[int, ...] | None = None,
    ) -> Path:
        import mss
        from PIL import Image

        target_dir = self.root_dir / self.run_id / subdir
        target_dir.mkdir(parents=True, exist_ok=True)
        safe_label = "".join(ch if ch.isalnum() or ch in ("-", "_") else "_" for ch in label)
        path = target_dir / f"{datetime.now():%H%M%S_%f}_{safe_label}.png"

        with mss.mss() as sct:
            monitor = normalize_region(region) or sct.monitors[1]
            shot = sct.grab(monitor)
            image = Image.frombytes("RGB", shot.size, shot.rgb)
            image.save(path)
        return path

    def annotate(self, image_path: str | Path, boxes: list[OCRBox], output_path: str | Path | None = None) -> Path:
        from PIL import Image, ImageDraw

        image_path = Path(image_path)
        output_path = Path(output_path) if output_path else image_path.with_name(f"{image_path.stem}_marked.png")
        output_path.parent.mkdir(parents=True, exist_ok=True)

        image = Image.open(image_path).convert("RGB")
        draw = ImageDraw.Draw(image)
        for box in boxes:
            draw.rectangle((box.x1, box.y1, box.x2, box.y2), outline="red", width=3)
            draw.ellipse((box.center_x - 5, box.center_y - 5, box.center_x + 5, box.center_y + 5), fill="red")
            draw.text((box.x1, max(0, box.y1 - 14)), box.text, fill="red")
        image.save(output_path)
        return output_path
