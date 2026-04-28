# office-ocr-rpa-template

A reusable office automation template based on screen capture, OCR text
recognition, coordinate clicking, and data-driven workflow loops.

The project is intentionally small. It reuses mature open-source components for
the hard parts, then keeps the workflow layer configurable with YAML.

## Plan

- Use RapidOCR as the default offline OCR engine.
- Use mss and Pillow for screenshots and target annotation.
- Use pyautogui for mouse, keyboard, hotkey, and scroll operations.
- Use openpyxl for Excel result output.
- Use YAML files for repeatable workflow templates.
- Save logs, OCR output, screenshots, and per-row status after each item.

Related projects reviewed before implementation:

- RapidOCR: https://github.com/RapidAI/RapidOCR
- PaddleOCR: https://github.com/PaddlePaddle/PaddleOCR
- RPA for Python: https://github.com/tebelorg/RPA-Python
- Robocorp RPA Framework: https://github.com/robocorp/rpaframework
- Umi-OCR: https://github.com/hiroi-sora/Umi-OCR

## Setup

Install Python 3.10 or newer. Python 3.11/3.12 is recommended.

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -U pip
pip install -r requirements.txt
```

RapidOCR may download model files on first use.

## Quick Start

By default, commands run in dry-run mode: capture the screen, run OCR, mark the
target, but do not click.

```powershell
python main.py click "Submit"
```

Real click:

```powershell
python main.py click "Submit" --live
```

Run the sample data-driven workflow:

```powershell
python main.py run tasks/batch_names.yaml
```

Run it with real mouse and keyboard actions:

```powershell
python main.py run tasks/batch_names.yaml --live
```

Run only rows with specific statuses:

```powershell
python main.py run tasks/batch_names.yaml --only "failed,pending"
```

For Chinese UI text, pass the real text directly from your terminal or put it in
the YAML file. The sample workflow uses YAML Unicode escapes for Chinese button
labels so the file remains codepage-safe.

## Data Format

CSV, TXT, and XLSX inputs are supported. Recommended columns:

```csv
task_id,name,keyword,status,remark
001,demo_name_1,demo_name_1,pending,
002,demo_name_2,demo_name_2,pending,
```

`keyword` is the text to find on screen. If it is empty, the runner falls back to
`name`.

## Workflow Actions

The MVP supports:

- `screenshot_ocr`
- `click_text`
- `wait_text`
- `input_text`
- `click_xy`
- `hotkey`
- `scroll`
- `if_text_exists`
- `update_status`

Example:

```yaml
steps:
  - name: click current name
    action: click_text
    text: "{{name}}"
    retry: 3
    wait_after: 1

  - name: click confirm
    action: click_text
    text: "\u786e\u8ba4"
    occurrence: 1
```

## Outputs

Each run creates a run id and writes:

```text
logs/<run_id>/run.log
logs/<run_id>/events.jsonl
logs/<run_id>/ocr.jsonl
logs/<run_id>/result.xlsx
screenshots/<run_id>/before/
screenshots/<run_id>/error/
```

## Safety

- Default mode is dry-run.
- Use `--live` for real mouse and keyboard actions.
- A startup countdown runs before real clicking.
- `Esc` attempts to stop the workflow.
- `pyautogui.FAILSAFE` is enabled by default: move the mouse to the upper-left
  screen corner to stop.
- The result workbook is saved after each row.

## Tests

```powershell
pytest
```

The tests use fake OCR data and do not move the mouse.
