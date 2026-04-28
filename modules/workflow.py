from __future__ import annotations

import re
from datetime import datetime
from pathlib import Path
from typing import Any

from .clicker import Clicker, ClickOptions
from .data_loader import load_task_table
from .logger import AutomationLogger
from .matcher import TextQuery, find_text
from .models import ClickTarget, RuntimeContext
from .ocr import RapidOCREngine
from .safety import StopController, StopRequested
from .screenshot import ScreenshotManager


class StepFailed(RuntimeError):
    pass


STATUS_PROCESSING = "\u5904\u7406\u4e2d"
STATUS_DONE = "\u5df2\u5b8c\u6210"
STATUS_FAILED = "\u5931\u8d25"
STATUS_STOPPED = "\u4e2d\u65ad"
RESULT_SUCCESS = "\u6210\u529f"
REMARK_STOPPED = "\u7528\u6237\u4e2d\u65ad"


def render_template(value: Any, item: dict[str, Any] | None) -> Any:
    if not isinstance(value, str) or item is None:
        return value

    def repl(match: re.Match[str]) -> str:
        key = match.group(1).strip()
        return str(item.get(key, ""))

    return re.sub(r"\{\{\s*([^}]+)\s*\}\}", repl, value)


def render_step(step: dict[str, Any], item: dict[str, Any] | None) -> dict[str, Any]:
    rendered: dict[str, Any] = {}
    for key, value in step.items():
        if isinstance(value, list):
            rendered[key] = [render_step(v, item) if isinstance(v, dict) else render_template(v, item) for v in value]
        elif isinstance(value, dict):
            rendered[key] = render_step(value, item)
        else:
            rendered[key] = render_template(value, item)
    return rendered


class WorkflowRunner:
    def __init__(
        self,
        config: dict[str, Any],
        ocr: RapidOCREngine,
        screenshot: ScreenshotManager,
        clicker: Clicker,
        logger: AutomationLogger,
        stop: StopController,
        context: RuntimeContext,
    ) -> None:
        self.config = config
        self.ocr_engine = ocr
        self.screenshot = screenshot
        self.clicker = clicker
        self.logger = logger
        self.stop = stop
        self.context = context
        self.poll_interval = float(config.get("workflow", {}).get("poll_interval", 0.5))

    def screenshot_ocr(self, label: str, subdir: str = "before", region: Any = None) -> list:
        image_path = self.screenshot.capture(label=label, subdir=subdir, region=region)
        boxes = self.ocr_engine.recognize(image_path)
        self.context.last_screenshot = image_path
        self.context.last_ocr = boxes
        self.logger.ocr(image_path, boxes)
        return boxes

    def _current_ocr(self, step_name: str, region: Any = None, force: bool = False) -> list:
        if force or not self.context.last_ocr:
            return self.screenshot_ocr(step_name, region=region)
        return self.context.last_ocr

    def _find(self, step: dict[str, Any]) -> Any:
        boxes = self._current_ocr(step.get("name", "find_text"), region=step.get("region"), force=step.get("refresh", False))
        query = TextQuery(
            text=str(step["text"]),
            match=step.get("match", "contains"),
            min_confidence=float(step.get("min_confidence", self.config.get("ocr", {}).get("min_confidence", 0.0))),
            fuzzy_threshold=int(step.get("fuzzy_threshold", 80)),
            occurrence=int(step.get("occurrence", 1)),
            position=step.get("position", "reading"),
            region=step.get("region"),
        )
        return find_text(boxes, query)

    def click_text(self, step: dict[str, Any]) -> ClickTarget:
        retry = int(step.get("retry", self.config.get("click", {}).get("retry", 1)))
        retry_interval = float(step.get("retry_interval", self.config.get("click", {}).get("retry_interval", 1)))
        last_error = ""
        for attempt in range(1, retry + 1):
            self.stop.check()
            box = self._find({**step, "refresh": True})
            if box:
                options = ClickOptions(
                    button=step.get("button", "left"),
                    clicks=int(step.get("clicks", 1)),
                    offset_x=int(step.get("offset_x", 0)),
                    offset_y=int(step.get("offset_y", 0)),
                    move_duration=float(step.get("move_duration", self.config.get("click", {}).get("move_duration", 0.15))),
                    dry_run=self.context.dry_run,
                )
                target = ClickTarget(box.center_x, box.center_y, box=box, description=f"text={step['text']}")
                clicked = self.clicker.click(target, options)
                wait_after = float(step.get("wait_after", 0))
                if wait_after:
                    self.stop.sleep(wait_after)
                return clicked
            last_error = f"Text not found: {step['text']} (attempt {attempt}/{retry})"
            if attempt < retry:
                self.stop.sleep(retry_interval)
        raise StepFailed(last_error)

    def wait_text(self, step: dict[str, Any]) -> Any:
        timeout = float(step.get("timeout", 10))
        end_at = datetime.now().timestamp() + timeout
        while datetime.now().timestamp() <= end_at:
            self.stop.check()
            box = self._find({**step, "refresh": True})
            if box:
                return box
            self.stop.sleep(self.poll_interval)
        raise StepFailed(f"Timed out waiting for text: {step['text']}")

    def run_step(self, step: dict[str, Any]) -> None:
        action = step.get("action")
        name = step.get("name", action)
        item = self.context.current_item
        try:
            if action == "screenshot_ocr":
                self.screenshot_ocr(name, region=step.get("region"))
            elif action == "click_text":
                self.click_text(step)
            elif action == "wait_text":
                self.wait_text(step)
            elif action == "input_text":
                self.clicker.input_text(str(step.get("text", "")))
            elif action == "click_xy":
                self.clicker.click_xy(int(step["x"]), int(step["y"]), ClickOptions(dry_run=self.context.dry_run))
            elif action == "hotkey":
                self.clicker.hotkey(step.get("keys", []))
            elif action == "scroll":
                self.clicker.scroll(int(step.get("amount", 0)))
            elif action == "if_text_exists":
                branch = step.get("then", []) if self._find(step) else step.get("else", [])
                for branch_step in branch:
                    self.run_step(render_step(branch_step, item))
            elif action == "update_status":
                if item is not None:
                    item["status"] = step.get("status", item.get("status", ""))
                    item["remark"] = step.get("remark", item.get("remark", ""))
            else:
                raise StepFailed(f"Unsupported action: {action}")
            self.logger.step(item, name, RESULT_SUCCESS)
        except Exception as exc:
            self.logger.step(item, name, STATUS_FAILED, str(exc))
            raise

    def run_template(
        self,
        template: dict[str, Any],
        only_statuses: set[str] | None = None,
        result_path: str | Path | None = None,
    ) -> Path:
        table = load_task_table(template["data_source"])
        skip_statuses = set(self.config.get("workflow", {}).get("skip_statuses", []))
        result_path = Path(result_path or Path(self.config.get("paths", {}).get("logs", "logs")) / self.context.run_id / "result.xlsx")
        steps = template.get("steps", [])

        for item in table.selectable_rows(only_statuses=only_statuses, skip_statuses=skip_statuses):
            self.context.current_item = item
            table.update(item, status=STATUS_PROCESSING, remark="")
            table.save_xlsx(result_path)
            try:
                for raw_step in steps:
                    self.stop.check()
                    self.run_step(render_step(raw_step, item))
                if item.get("status") == STATUS_PROCESSING:
                    table.update(item, status=STATUS_DONE, remark=RESULT_SUCCESS)
            except StopRequested:
                table.update(item, status=STATUS_STOPPED, remark=REMARK_STOPPED)
                table.save_xlsx(result_path)
                raise
            except Exception as exc:
                table.update(item, status=STATUS_FAILED, remark=str(exc))
                self._save_error_screenshot(item)
                if not self.config.get("workflow", {}).get("continue_on_item_error", True):
                    table.save_xlsx(result_path)
                    raise
            finally:
                table.save_xlsx(result_path)
        return result_path

    def _save_error_screenshot(self, item: dict[str, Any]) -> None:
        label = str(item.get("name") or item.get("task_id") or "error")
        try:
            self.screenshot.capture(label=label, subdir="error")
        except Exception as exc:
            self.logger.event("error_screenshot_failed", {"error": str(exc), "item": label})
