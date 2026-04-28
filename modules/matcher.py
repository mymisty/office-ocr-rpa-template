from __future__ import annotations

from dataclasses import dataclass
from typing import Any

from rapidfuzz import fuzz

from .models import OCRBox


@dataclass(slots=True)
class TextQuery:
    text: str
    match: str = "contains"
    min_confidence: float = 0.0
    fuzzy_threshold: int = 80
    occurrence: int = 1
    position: str = "reading"
    region: dict[str, int] | None = None


def _in_region(box: OCRBox, region: dict[str, int] | None) -> bool:
    if not region:
        return True
    x1 = int(region.get("x1", region.get("left", 0)))
    y1 = int(region.get("y1", region.get("top", 0)))
    x2 = int(region.get("x2", x1 + int(region.get("width", 0))))
    y2 = int(region.get("y2", y1 + int(region.get("height", 0))))
    return box.center_x >= x1 and box.center_x <= x2 and box.center_y >= y1 and box.center_y <= y2


def _matches(box_text: str, query: TextQuery) -> bool:
    target = query.text.strip()
    actual = box_text.strip()
    if not target:
        return False
    if query.match == "exact":
        return actual == target
    if query.match == "contains":
        return target in actual or actual in target
    if query.match == "fuzzy":
        return fuzz.ratio(actual, target) >= query.fuzzy_threshold
    raise ValueError(f"Unknown text match mode: {query.match}")


def _sort_boxes(boxes: list[OCRBox], position: str) -> list[OCRBox]:
    if position == "top":
        return sorted(boxes, key=lambda b: (b.y1, b.x1))
    if position == "bottom":
        return sorted(boxes, key=lambda b: (-b.y2, b.x1))
    if position == "left":
        return sorted(boxes, key=lambda b: (b.x1, b.y1))
    if position == "right":
        return sorted(boxes, key=lambda b: (-b.x2, b.y1))
    return sorted(boxes, key=lambda b: (b.y1, b.x1))


def find_text(boxes: list[OCRBox], query: TextQuery | dict[str, Any] | str) -> OCRBox | None:
    if isinstance(query, str):
        query = TextQuery(text=query)
    elif isinstance(query, dict):
        query = TextQuery(
            text=str(query["text"]),
            match=query.get("match", "contains"),
            min_confidence=float(query.get("min_confidence", 0.0)),
            fuzzy_threshold=int(query.get("fuzzy_threshold", 80)),
            occurrence=int(query.get("occurrence", 1)),
            position=query.get("position", "reading"),
            region=query.get("region"),
        )

    candidates = [
        box
        for box in boxes
        if box.confidence >= query.min_confidence and _in_region(box, query.region) and _matches(box.text, query)
    ]
    candidates = _sort_boxes(candidates, query.position)
    index = max(1, query.occurrence) - 1
    return candidates[index] if index < len(candidates) else None
