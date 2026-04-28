from __future__ import annotations

import json
from pathlib import Path
from typing import Any, Iterable

from .models import OCRBox


def _flatten_points(points: Any) -> list[tuple[int, int]]:
    if points is None:
        return []
    if isinstance(points, dict):
        if {"x1", "y1", "x2", "y2"} <= set(points):
            return [(int(points["x1"]), int(points["y1"])), (int(points["x2"]), int(points["y2"]))]
        if {"left", "top", "width", "height"} <= set(points):
            left = int(points["left"])
            top = int(points["top"])
            return [(left, top), (left + int(points["width"]), top + int(points["height"]))]
    if isinstance(points, (list, tuple)):
        pairs: list[tuple[int, int]] = []
        for item in points:
            if isinstance(item, (list, tuple)) and len(item) >= 2:
                pairs.append((int(float(item[0])), int(float(item[1]))))
        return pairs
    return []


def _box_from_points(text: str, points: Any, confidence: float, source_image: str) -> OCRBox | None:
    pairs = _flatten_points(points)
    if not pairs:
        return None
    xs = [p[0] for p in pairs]
    ys = [p[1] for p in pairs]
    return OCRBox(
        text=str(text).strip(),
        x1=min(xs),
        y1=min(ys),
        x2=max(xs),
        y2=max(ys),
        confidence=float(confidence),
        source_image=source_image,
    )


def _iter_from_result(result: Any) -> Iterable[tuple[Any, str, float]]:
    if result is None:
        return []

    if hasattr(result, "to_json"):
        try:
            raw = result.to_json()
            payload = json.loads(raw) if isinstance(raw, str) else raw
            items = payload.get("data") or payload.get("res") or payload.get("result") or []
            parsed = []
            for item in items:
                text = item.get("text") or item.get("rec_text") or item.get("txt") or ""
                score = item.get("confidence") or item.get("score") or item.get("rec_score") or 1.0
                points = item.get("box") or item.get("points") or item.get("dt_polys")
                parsed.append((points, text, score))
            if parsed:
                return parsed
        except Exception:
            pass

    boxes = getattr(result, "boxes", None) or getattr(result, "dt_polys", None)
    texts = getattr(result, "txts", None) or getattr(result, "texts", None) or getattr(result, "rec_texts", None)
    scores = getattr(result, "scores", None) or getattr(result, "rec_scores", None)
    if boxes is not None and texts is not None:
        return zip(boxes, texts, scores or [1.0] * len(texts))

    if isinstance(result, dict):
        items = result.get("data") or result.get("res") or result.get("result") or result.get("ocr_result")
        if items is not None:
            parsed = []
            for item in items:
                text = item.get("text") or item.get("rec_text") or item.get("txt") or ""
                score = item.get("confidence") or item.get("score") or item.get("rec_score") or 1.0
                points = item.get("box") or item.get("points") or item.get("dt_polys")
                parsed.append((points, text, score))
            return parsed

    if isinstance(result, (list, tuple)):
        parsed = []
        for item in result:
            if isinstance(item, dict):
                text = item.get("text") or item.get("rec_text") or item.get("txt") or ""
                score = item.get("confidence") or item.get("score") or item.get("rec_score") or 1.0
                points = item.get("box") or item.get("points") or item.get("dt_polys")
                parsed.append((points, text, score))
            elif isinstance(item, (list, tuple)) and len(item) >= 3:
                parsed.append((item[0], item[1], item[2]))
            elif isinstance(item, (list, tuple)) and len(item) == 2:
                points, rec = item
                if isinstance(rec, (list, tuple)) and len(rec) >= 2:
                    parsed.append((points, rec[0], rec[1]))
        return parsed

    return []


class RapidOCREngine:
    def __init__(self, min_confidence: float = 0.0) -> None:
        self.min_confidence = min_confidence
        self._engine = None

    def _load(self) -> Any:
        if self._engine is None:
            from rapidocr import RapidOCR

            self._engine = RapidOCR()
        return self._engine

    def recognize(self, image_path: str | Path) -> list[OCRBox]:
        image_path = Path(image_path)
        result = self._load()(str(image_path))
        boxes: list[OCRBox] = []
        for points, text, score in _iter_from_result(result):
            if not str(text).strip():
                continue
            box = _box_from_points(str(text), points, float(score), str(image_path))
            if box and box.confidence >= self.min_confidence:
                boxes.append(box)
        return boxes


def create_ocr_engine(config: dict[str, Any]) -> RapidOCREngine:
    ocr_config = config.get("ocr", {})
    engine = ocr_config.get("engine", "rapidocr").lower()
    if engine != "rapidocr":
        raise ValueError(f"Unsupported OCR engine for this MVP: {engine}")
    return RapidOCREngine(min_confidence=float(ocr_config.get("min_confidence", 0.0)))
