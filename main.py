from __future__ import annotations

import argparse
import sys
from datetime import datetime
from pathlib import Path
from typing import Any

from modules.clicker import Clicker, ClickOptions
from modules.config import load_config, load_yaml
from modules.logger import AutomationLogger
from modules.matcher import TextQuery, find_text
from modules.models import ClickTarget, RuntimeContext
from modules.ocr import create_ocr_engine
from modules.safety import StopController, StopRequested
from modules.screenshot import ScreenshotManager
from modules.workflow import WorkflowRunner


RESULT_SUCCESS = "\u6210\u529f"
RESULT_FAILED = "\u5931\u8d25"
REMARK_TEXT_NOT_FOUND = "\u672a\u627e\u5230\u76ee\u6807\u6587\u5b57"
HELP_ONLY_STATUS_EXAMPLE = "\u5931\u8d25,\u5f85\u5904\u7406"


def parse_statuses(value: str | None) -> set[str] | None:
    if not value:
        return None
    return {item.strip() for item in value.split(",") if item.strip()}


def make_run_id(prefix: str = "run") -> str:
    return f"{prefix}_{datetime.now():%Y%m%d_%H%M%S}"


def resolve_dry_run(args: argparse.Namespace, config: dict[str, Any]) -> bool:
    if getattr(args, "live", False):
        return False
    if getattr(args, "dry_run", False):
        return True
    return bool(config.get("app", {}).get("default_dry_run", True))


def countdown(seconds: int, stop: StopController) -> None:
    if seconds <= 0:
        return
    print(f"Automation starts in {seconds} seconds. Press Esc to stop.")
    for left in range(seconds, 0, -1):
        print(f"{left}...")
        stop.sleep(1)


def build_runtime(config: dict[str, Any], run_id: str, dry_run: bool) -> tuple:
    screenshots_root = Path(config.get("paths", {}).get("screenshots", "screenshots"))
    logs_root = Path(config.get("paths", {}).get("logs", "logs"))
    stop = StopController(hotkey=config.get("app", {}).get("stop_hotkey", "esc"))
    stop.start()
    context = RuntimeContext(run_id=run_id, root_dir=Path.cwd(), dry_run=dry_run)
    screenshot = ScreenshotManager(screenshots_root, run_id)
    logger = AutomationLogger(logs_root, run_id)
    ocr = create_ocr_engine(config)
    clicker = Clicker(
        dry_run=dry_run,
        fail_safe=bool(config.get("app", {}).get("fail_safe_corner", True)),
        stop=stop,
    )
    return context, stop, screenshot, logger, ocr, clicker


def command_click(args: argparse.Namespace) -> int:
    config = load_config(args.config)
    dry_run = resolve_dry_run(args, config)
    run_id = args.run_id or make_run_id("click")
    context, stop, screenshot, logger, ocr, clicker = build_runtime(config, run_id, dry_run)
    try:
        if not dry_run:
            countdown(int(config.get("app", {}).get("startup_countdown_seconds", 3)), stop)
        image_path = screenshot.capture(label="click_before", subdir="before", region=args.region)
        boxes = ocr.recognize(image_path, screen_offset=screenshot.screen_offset_for(image_path))
        logger.ocr(image_path, boxes)
        query = TextQuery(
            text=args.text,
            match=args.match,
            min_confidence=float(args.min_confidence or config.get("ocr", {}).get("min_confidence", 0.0)),
            fuzzy_threshold=args.fuzzy_threshold,
            occurrence=args.occurrence,
            position=args.position,
        )
        box = find_text(boxes, query)
        if not box:
            screenshot.capture(label="click_not_found", subdir="error", region=args.region)
            logger.event("click_text", {"text": args.text, "result": RESULT_FAILED, "remark": REMARK_TEXT_NOT_FOUND})
            print(f"Text not found: {args.text}")
            return 2

        marked = screenshot.annotate(image_path, [box])
        target = ClickTarget(box.screen_center_x, box.screen_center_y, box=box, description=f"text={args.text}")
        clicked = clicker.click(
            target,
            ClickOptions(
                offset_x=args.offset_x,
                offset_y=args.offset_y,
                clicks=args.clicks,
                button=args.button,
                dry_run=dry_run,
            ),
        )
        logger.event(
            "click_text",
            {
                "text": args.text,
                "result": RESULT_SUCCESS,
                "x": clicked.x,
                "y": clicked.y,
                "dry_run": dry_run,
                "marked_image": str(marked),
            },
        )
        print(f"Matched '{box.text}' at ({clicked.x}, {clicked.y}). Marked image: {marked}")
        return 0
    except StopRequested:
        logger.event("stopped", {"reason": "user"})
        print("Stopped by user.")
        return 130
    finally:
        stop.close()


def command_run(args: argparse.Namespace) -> int:
    config = load_config(args.config)
    template = load_yaml(args.template)
    dry_run = resolve_dry_run(args, config)
    run_id = args.run_id or make_run_id("run")
    context, stop, screenshot, logger, ocr, clicker = build_runtime(config, run_id, dry_run)
    runner = WorkflowRunner(config, ocr, screenshot, clicker, logger, stop, context)
    try:
        if not dry_run:
            countdown(int(config.get("app", {}).get("startup_countdown_seconds", 3)), stop)
        result_path = runner.run_template(
            template,
            only_statuses=parse_statuses(args.only),
            result_path=args.result,
        )
        print(f"Workflow finished. Result: {result_path}")
        return 0
    except StopRequested:
        print("Stopped by user.")
        return 130
    finally:
        stop.close()


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Office OCR RPA template runner")
    parser.add_argument("--config", default="config.yaml", help="Global config YAML path")
    parser.add_argument("--run-id", default=None, help="Custom run id for logs and screenshots")

    subparsers = parser.add_subparsers(dest="command", required=True)

    click_parser = subparsers.add_parser("click", help="Find text on screen and click it")
    click_parser.add_argument("text", help="Target text to find")
    click_safety = click_parser.add_mutually_exclusive_group()
    click_safety.add_argument("--dry-run", action="store_true", help="Recognize and mark target without clicking")
    click_safety.add_argument("--live", action="store_true", help="Allow real mouse/keyboard actions")
    click_parser.add_argument("--match", choices=["exact", "contains", "fuzzy"], default="contains")
    click_parser.add_argument("--fuzzy-threshold", type=int, default=80)
    click_parser.add_argument("--min-confidence", type=float, default=None)
    click_parser.add_argument("--occurrence", type=int, default=1, help="1-based match index")
    click_parser.add_argument("--position", choices=["reading", "top", "bottom", "left", "right"], default="reading")
    click_parser.add_argument("--offset-x", type=int, default=0)
    click_parser.add_argument("--offset-y", type=int, default=0)
    click_parser.add_argument("--clicks", type=int, default=1)
    click_parser.add_argument("--button", choices=["left", "right", "middle"], default="left")
    click_parser.add_argument("--region", nargs=4, type=int, metavar=("X1", "Y1", "X2", "Y2"), default=None)
    click_parser.set_defaults(func=command_click)

    run_parser = subparsers.add_parser("run", help="Run a YAML workflow")
    run_parser.add_argument("template", help="Workflow YAML path")
    run_safety = run_parser.add_mutually_exclusive_group()
    run_safety.add_argument("--dry-run", action="store_true", help="Run recognition and logs without clicking/typing")
    run_safety.add_argument("--live", action="store_true", help="Allow real mouse/keyboard actions")
    run_parser.add_argument(
        "--only",
        default=None,
        help=f"Comma-separated statuses to process, for example: {HELP_ONLY_STATUS_EXAMPLE}",
    )
    run_parser.add_argument("--result", default=None, help="Result XLSX path")
    run_parser.set_defaults(func=command_run)

    return parser


def main() -> int:
    parser = build_parser()
    args = parser.parse_args()
    return args.func(args)


if __name__ == "__main__":
    sys.exit(main())
