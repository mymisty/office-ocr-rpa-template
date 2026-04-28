from __future__ import annotations

import json
import logging
from datetime import datetime
from pathlib import Path
from typing import Any

from .models import OCRBox


class AutomationLogger:
    def __init__(self, root_dir: str | Path, run_id: str) -> None:
        self.root_dir = Path(root_dir)
        self.run_id = run_id
        self.run_dir = self.root_dir / run_id
        self.run_dir.mkdir(parents=True, exist_ok=True)
        self.log_path = self.run_dir / "run.log"
        self.events_path = self.run_dir / "events.jsonl"
        self.ocr_path = self.run_dir / "ocr.jsonl"

        self.logger = logging.getLogger(f"office_auto.{run_id}")
        self.logger.setLevel(logging.INFO)
        self.logger.handlers.clear()
        handler = logging.FileHandler(self.log_path, encoding="utf-8")
        handler.setFormatter(logging.Formatter("%(asctime)s %(levelname)s %(message)s"))
        self.logger.addHandler(handler)

    def event(self, event_type: str, payload: dict[str, Any]) -> None:
        entry = {"time": datetime.now().isoformat(timespec="seconds"), "type": event_type, **payload}
        with self.events_path.open("a", encoding="utf-8") as handle:
            handle.write(json.dumps(entry, ensure_ascii=False) + "\n")
        self.logger.info("%s %s", event_type, json.dumps(payload, ensure_ascii=False))

    def step(self, item: dict[str, Any] | None, step_name: str, result: str, remark: str = "") -> None:
        self.event(
            "step",
            {
                "task_id": item.get("task_id") if item else "",
                "name": item.get("name") if item else "",
                "step": step_name,
                "result": result,
                "remark": remark,
            },
        )

    def ocr(self, image_path: str | Path, boxes: list[OCRBox]) -> None:
        payload = {"image": str(image_path), "boxes": [box.to_dict() for box in boxes]}
        with self.ocr_path.open("a", encoding="utf-8") as handle:
            handle.write(json.dumps(payload, ensure_ascii=False) + "\n")
        self.event("ocr", {"image": str(image_path), "count": len(boxes)})
