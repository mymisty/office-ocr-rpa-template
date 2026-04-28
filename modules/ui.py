from __future__ import annotations

import ctypes
import queue
import subprocess
import sys
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from typing import Callable

import yaml


MATCH_MODES = {
    "包含": "contains",
    "精确": "exact",
    "模糊": "fuzzy",
}

ACTION_SNIPPETS = {
    "截图OCR": """  - name: screenshot and OCR
    action: screenshot_ocr
""",
    "点击文字": """  - name: click target text
    action: click_text
    text: "{{keyword}}"
    retry: 3
    wait_after: 1
""",
    "等待文字": """  - name: wait for text
    action: wait_text
    text: "详情"
    timeout: 5
""",
    "输入文字": """  - name: input text
    action: input_text
    text: "{{name}}"
""",
    "快捷键": """  - name: press hotkey
    action: hotkey
    keys: ["ctrl", "s"]
""",
    "条件分支": """  - name: branch by text
    action: if_text_exists
    text: "确认"
    then:
      - action: click_text
        text: "确认"
    else:
      - action: scroll
        amount: -5
""",
    "更新状态": """  - name: mark current item done
    action: update_status
    status: "已完成"
    remark: "成功"
""",
}


def enable_dpi_awareness() -> None:
    if not hasattr(ctypes, "windll"):
        return
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(2)
    except Exception:
        try:
            ctypes.windll.user32.SetProcessDPIAware()
        except Exception:
            pass


class FluentRpaApp:
    def __init__(self, root: tk.Tk, default_config: str = "config.yaml") -> None:
        self.root = root
        self.root_dir = Path.cwd()
        self.main_path = self.root_dir / "main.py"
        self.output_queue: queue.Queue[str] = queue.Queue()
        self.process: subprocess.Popen[str] | None = None

        self.config_var = tk.StringVar(value=default_config)
        self.run_id_var = tk.StringVar()
        self.live_var = tk.BooleanVar(value=False)
        self.template_status_var = tk.StringVar(value="模板未加载")

        self.click_text_var = tk.StringVar(value="提交")
        self.match_var = tk.StringVar(value="包含")
        self.confidence_var = tk.StringVar(value="")
        self.occurrence_var = tk.StringVar(value="1")
        self.region_var = tk.StringVar()

        self.template_var = tk.StringVar(value="tasks/batch_names.yaml")
        self.only_var = tk.StringVar()
        self.result_var = tk.StringVar()

        self._configure_window()
        self._configure_styles()
        self._build_layout()
        self._load_template_to_editor(silent=True)
        self._poll_output()

    def _configure_window(self) -> None:
        self.root.title("Office OCR RPA 控制台")
        self.root.geometry("1360x860")
        self.root.minsize(1180, 760)
        self.root.configure(bg="#f3f3f3")
        try:
            self.root.tk.call("tk", "scaling", 1.2)
        except tk.TclError:
            pass

    def _configure_styles(self) -> None:
        self.style = ttk.Style(self.root)
        self.style.theme_use("clam")
        self.style.configure(".", font=("Microsoft YaHei UI", 10), background="#f3f3f3", foreground="#1f1f1f")
        self.style.configure("App.TFrame", background="#f3f3f3")
        self.style.configure("Surface.TFrame", background="#ffffff", relief="flat")
        self.style.configure("Sidebar.TFrame", background="#fbfbfb")
        self.style.configure("Title.TLabel", font=("Microsoft YaHei UI", 24, "bold"), background="#f3f3f3")
        self.style.configure("Subtitle.TLabel", font=("Microsoft YaHei UI", 10), background="#f3f3f3", foreground="#606060")
        self.style.configure("CardTitle.TLabel", font=("Microsoft YaHei UI", 12, "bold"), background="#ffffff")
        self.style.configure("CardText.TLabel", background="#ffffff", foreground="#4d4d4d")
        self.style.configure("TLabel", background="#ffffff")
        self.style.configure("TCheckbutton", background="#ffffff")
        self.style.configure("TEntry", fieldbackground="#ffffff", bordercolor="#d0d0d0", lightcolor="#d0d0d0", padding=4)
        self.style.configure("TCombobox", fieldbackground="#ffffff", bordercolor="#d0d0d0", padding=4)
        self.style.configure("Accent.TButton", background="#0078d4", foreground="#ffffff", bordercolor="#0078d4", padding=(14, 7))
        self.style.map("Accent.TButton", background=[("active", "#106ebe"), ("disabled", "#c8c8c8")])
        self.style.configure("Ghost.TButton", background="#ffffff", bordercolor="#d0d0d0", padding=(12, 7))
        self.style.map("Ghost.TButton", background=[("active", "#f3f3f3")])
        self.style.configure("Nav.TButton", anchor="w", background="#fbfbfb", bordercolor="#fbfbfb", padding=(16, 11))
        self.style.map("Nav.TButton", background=[("active", "#eef6fc")])

    def _build_layout(self) -> None:
        shell = ttk.Frame(self.root, style="App.TFrame", padding=18)
        shell.pack(fill=tk.BOTH, expand=True)
        shell.columnconfigure(1, weight=1)
        shell.rowconfigure(0, weight=1)

        sidebar = ttk.Frame(shell, style="Sidebar.TFrame", padding=(10, 16))
        sidebar.grid(row=0, column=0, sticky="nsew", padx=(0, 16))
        sidebar.configure(width=232)
        sidebar.grid_propagate(False)

        brand = ttk.Label(sidebar, text="Office OCR", font=("Microsoft YaHei UI", 19, "bold"), background="#fbfbfb")
        brand.pack(anchor="w", padx=8, pady=(0, 22))
        ttk.Button(sidebar, text="快速点击", style="Nav.TButton", command=self._focus_click).pack(fill=tk.X, pady=2)
        ttk.Button(sidebar, text="批量流程", style="Nav.TButton", command=self._focus_workflow).pack(fill=tk.X, pady=2)
        ttk.Button(sidebar, text="模板编辑", style="Nav.TButton", command=self._focus_template).pack(fill=tk.X, pady=2)
        ttk.Button(sidebar, text="运行日志", style="Nav.TButton", command=self._focus_log).pack(fill=tk.X, pady=2)
        ttk.Frame(sidebar, style="Sidebar.TFrame").pack(fill=tk.BOTH, expand=True)
        ttk.Label(sidebar, textvariable=self._mode_label(), background="#fbfbfb", foreground="#606060").pack(anchor="w", padx=8)

        content = ttk.Frame(shell, style="App.TFrame")
        content.grid(row=0, column=1, sticky="nsew")
        content.columnconfigure(0, weight=1)
        content.rowconfigure(2, weight=1)

        header = ttk.Frame(content, style="App.TFrame")
        header.grid(row=0, column=0, sticky="ew", pady=(0, 16))
        header.columnconfigure(0, weight=1)
        ttk.Label(header, text="自动化控制台", style="Title.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(header, text="屏幕 OCR、批量流程、模板编辑与执行日志", style="Subtitle.TLabel").grid(
            row=1,
            column=0,
            sticky="w",
            pady=(3, 0),
        )

        self.global_card = self._card(content, row=1, title="运行设置")
        self.global_card.columnconfigure(1, weight=1)
        self._field(self.global_card, "配置文件", self.config_var, 0, browse=lambda: self._browse_file(self.config_var, "YAML", "*.yaml"))
        self._field(self.global_card, "运行编号", self.run_id_var, 1)
        ttk.Checkbutton(self.global_card, text="真实执行鼠标和键盘操作", variable=self.live_var).grid(row=2, column=1, sticky="w", pady=(8, 0))

        work_area = ttk.Frame(content, style="App.TFrame")
        work_area.grid(row=2, column=0, sticky="nsew")
        work_area.columnconfigure(0, weight=1)
        work_area.columnconfigure(1, weight=1)
        work_area.rowconfigure(1, weight=3)
        work_area.rowconfigure(2, weight=2)

        self.click_card = self._card(work_area, row=0, column=0, title="快速点击", padx=(0, 8))
        self._field(self.click_card, "目标文字", self.click_text_var, 0)
        self._combo(self.click_card, "匹配方式", self.match_var, tuple(MATCH_MODES), 1)
        self._field(self.click_card, "最低置信度", self.confidence_var, 2)
        self._field(self.click_card, "匹配序号", self.occurrence_var, 3)
        self._field(self.click_card, "截图区域", self.region_var, 4)
        self._button_row(self.click_card, 5, ("执行点击", self.run_click, "Accent.TButton"), ("预演", self.preview_click, "Ghost.TButton"))

        self.workflow_card = self._card(work_area, row=0, column=1, title="批量流程", padx=(8, 0))
        self._field(
            self.workflow_card,
            "流程模板",
            self.template_var,
            0,
            browse=self._browse_template_file,
        )
        self._field(self.workflow_card, "只处理状态", self.only_var, 1)
        self._field(
            self.workflow_card,
            "结果文件",
            self.result_var,
            2,
            browse=lambda: self._browse_save_file(self.result_var),
        )
        self._button_row(
            self.workflow_card,
            3,
            ("运行流程", self.run_workflow, "Accent.TButton"),
            ("预演流程", self.preview_workflow, "Ghost.TButton"),
        )

        self.template_card = self._card(work_area, row=1, column=0, columnspan=2, title="流程模板编辑器", pady=(16, 0))
        self._build_template_editor(self.template_card)

        log_card = self._card(work_area, row=2, column=0, columnspan=2, title="运行日志", pady=(16, 0))
        self._build_log_view(log_card)

        footer = ttk.Frame(content, style="App.TFrame")
        footer.grid(row=3, column=0, sticky="ew", pady=(12, 0))
        footer.columnconfigure(0, weight=1)
        self.status_var = tk.StringVar(value="就绪")
        ttk.Label(footer, textvariable=self.status_var, style="Subtitle.TLabel").grid(row=0, column=0, sticky="w")
        self.stop_button = ttk.Button(footer, text="停止", style="Ghost.TButton", command=self.stop_process, state=tk.DISABLED)
        self.stop_button.grid(row=0, column=1, sticky="e")

    def _build_template_editor(self, parent: ttk.Frame) -> None:
        parent.rowconfigure(2, weight=1)
        parent.columnconfigure(0, weight=1)

        actions = ttk.Frame(parent, style="Surface.TFrame")
        actions.grid(row=1, column=0, sticky="ew", pady=(0, 10))
        actions.columnconfigure(8, weight=1)
        ttk.Button(actions, text="加载模板", style="Ghost.TButton", command=self._load_template_to_editor).grid(row=0, column=0, padx=(0, 8))
        ttk.Button(actions, text="保存模板", style="Accent.TButton", command=self._save_template_from_editor).grid(row=0, column=1, padx=(0, 8))
        ttk.Button(actions, text="另存为", style="Ghost.TButton", command=self._save_template_as).grid(row=0, column=2, padx=(0, 8))
        ttk.Button(actions, text="校验模板", style="Ghost.TButton", command=self._validate_template_editor).grid(row=0, column=3, padx=(0, 8))
        ttk.Button(actions, text="运行前预检", style="Ghost.TButton", command=self._preflight_check).grid(row=0, column=4, padx=(0, 8))
        ttk.Label(actions, textvariable=self.template_status_var, style="CardText.TLabel").grid(row=0, column=8, sticky="e")

        snippets = ttk.Frame(actions, style="Surface.TFrame")
        snippets.grid(row=1, column=0, columnspan=9, sticky="ew", pady=(10, 0))
        ttk.Label(snippets, text="插入动作", style="CardText.TLabel").pack(side=tk.LEFT, padx=(0, 8))
        for label in ACTION_SNIPPETS:
            ttk.Button(
                snippets,
                text=label,
                style="Ghost.TButton",
                command=lambda name=label: self._insert_template_snippet(name),
            ).pack(side=tk.LEFT, padx=(0, 6))

        editor_shell = ttk.Frame(parent, style="Surface.TFrame")
        editor_shell.grid(row=2, column=0, sticky="nsew")
        editor_shell.rowconfigure(0, weight=1)
        editor_shell.columnconfigure(0, weight=1)

        self.template_text = tk.Text(
            editor_shell,
            height=12,
            undo=True,
            wrap="none",
            relief="flat",
            bg="#fbfbfb",
            fg="#1f1f1f",
            insertbackground="#0078d4",
            selectbackground="#cfe8ff",
            font=("Cascadia Mono", 10),
            padx=12,
            pady=10,
        )
        self.template_text.grid(row=0, column=0, sticky="nsew")
        y_scroll = ttk.Scrollbar(editor_shell, command=self.template_text.yview)
        y_scroll.grid(row=0, column=1, sticky="ns")
        x_scroll = ttk.Scrollbar(editor_shell, orient=tk.HORIZONTAL, command=self.template_text.xview)
        x_scroll.grid(row=1, column=0, sticky="ew")
        self.template_text.configure(yscrollcommand=y_scroll.set, xscrollcommand=x_scroll.set)

    def _build_log_view(self, parent: ttk.Frame) -> None:
        parent.rowconfigure(2, weight=1)
        parent.columnconfigure(0, weight=1)
        self.command_preview = ttk.Label(parent, text="", style="CardText.TLabel")
        self.command_preview.grid(row=1, column=0, sticky="ew", pady=(0, 8))
        self.output = tk.Text(
            parent,
            height=8,
            wrap="word",
            relief="flat",
            bg="#1e1e1e",
            fg="#f5f5f5",
            insertbackground="#ffffff",
            font=("Cascadia Mono", 10),
            padx=12,
            pady=10,
        )
        self.output.grid(row=2, column=0, sticky="nsew")
        scrollbar = ttk.Scrollbar(parent, command=self.output.yview)
        scrollbar.grid(row=2, column=1, sticky="ns")
        self.output.configure(yscrollcommand=scrollbar.set)

    def _card(
        self,
        parent: ttk.Frame,
        row: int,
        title: str,
        column: int = 0,
        columnspan: int = 1,
        padx: tuple[int, int] = (0, 0),
        pady: tuple[int, int] = (0, 0),
    ) -> ttk.Frame:
        outer = ttk.Frame(parent, style="Surface.TFrame", padding=18)
        outer.grid(row=row, column=column, columnspan=columnspan, sticky="nsew", padx=padx, pady=pady)
        outer.columnconfigure(1, weight=1)
        ttk.Label(outer, text=title, style="CardTitle.TLabel").grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 12))
        return outer

    def _field(
        self,
        parent: ttk.Frame,
        label: str,
        variable: tk.StringVar,
        row: int,
        browse: Callable[[], None] | None = None,
    ) -> None:
        grid_row = row + 1
        ttk.Label(parent, text=label).grid(row=grid_row, column=0, sticky="w", pady=6, padx=(0, 10))
        entry = ttk.Entry(parent, textvariable=variable)
        entry.grid(row=grid_row, column=1, sticky="ew", pady=6)
        if browse:
            ttk.Button(parent, text="浏览", style="Ghost.TButton", command=browse).grid(row=grid_row, column=2, padx=(8, 0), pady=6)

    def _combo(self, parent: ttk.Frame, label: str, variable: tk.StringVar, values: tuple[str, ...], row: int) -> None:
        grid_row = row + 1
        ttk.Label(parent, text=label).grid(row=grid_row, column=0, sticky="w", pady=6, padx=(0, 10))
        combo = ttk.Combobox(parent, textvariable=variable, values=values, state="readonly")
        combo.grid(row=grid_row, column=1, sticky="ew", pady=6)

    def _button_row(self, parent: ttk.Frame, row: int, *buttons: tuple[str, Callable[[], None], str]) -> None:
        frame = ttk.Frame(parent, style="Surface.TFrame")
        frame.grid(row=row + 1, column=0, columnspan=3, sticky="e", pady=(14, 0))
        for text, command, style in buttons:
            ttk.Button(frame, text=text, command=command, style=style).pack(side=tk.LEFT, padx=(8, 0))

    def _mode_label(self) -> tk.StringVar:
        mode = tk.StringVar(value="默认：安全预演")

        def refresh(*_: object) -> None:
            mode.set("模式：真实执行" if self.live_var.get() else "默认：安全预演")

        self.live_var.trace_add("write", refresh)
        return mode

    def _browse_file(self, variable: tk.StringVar, label: str, pattern: str) -> None:
        path = filedialog.askopenfilename(initialdir=self.root_dir, filetypes=[(label, pattern), ("所有文件", "*.*")])
        if path:
            variable.set(self._display_path(Path(path)))

    def _browse_template_file(self) -> None:
        path = filedialog.askopenfilename(initialdir=self.root_dir, filetypes=[("YAML", "*.yaml"), ("YAML", "*.yml"), ("所有文件", "*.*")])
        if path:
            self.template_var.set(self._display_path(Path(path)))
            self._load_template_to_editor()

    def _browse_save_file(self, variable: tk.StringVar) -> None:
        path = filedialog.asksaveasfilename(
            initialdir=self.root_dir,
            defaultextension=".xlsx",
            filetypes=[("Excel 工作簿", "*.xlsx"), ("所有文件", "*.*")],
        )
        if path:
            variable.set(self._display_path(Path(path)))

    def _display_path(self, path: Path) -> str:
        try:
            return str(path.resolve().relative_to(self.root_dir))
        except ValueError:
            return str(path)

    def _resolve_path(self, value: str) -> Path:
        path = Path(value)
        if not path.is_absolute():
            path = self.root_dir / path
        return path

    def _load_template_to_editor(self, silent: bool = False) -> None:
        path_text = self.template_var.get().strip()
        if not path_text:
            if not silent:
                self.template_status_var.set("请先选择模板文件")
            return
        path = self._resolve_path(path_text)
        try:
            content = path.read_text(encoding="utf-8")
        except Exception as exc:
            self.template_status_var.set(f"加载失败：{exc}")
            if not silent:
                messagebox.showerror("加载模板失败", str(exc))
            return
        self.template_text.delete("1.0", tk.END)
        self.template_text.insert("1.0", content)
        self.template_text.edit_modified(False)
        self.template_status_var.set(f"已加载：{self._display_path(path)}")

    def _save_template_from_editor(self, silent: bool = False) -> bool:
        path_text = self.template_var.get().strip()
        if not path_text:
            self.template_status_var.set("请先选择模板文件")
            if not silent:
                messagebox.showwarning("保存模板", "请先选择模板文件。")
            return False
        path = self._resolve_path(path_text)
        try:
            path.parent.mkdir(parents=True, exist_ok=True)
            path.write_text(self.template_text.get("1.0", "end-1c"), encoding="utf-8")
        except Exception as exc:
            self.template_status_var.set(f"保存失败：{exc}")
            if not silent:
                messagebox.showerror("保存模板失败", str(exc))
            return False
        self.template_text.edit_modified(False)
        self.template_status_var.set(f"已保存：{self._display_path(path)}")
        return True

    def _save_template_as(self) -> None:
        path = filedialog.asksaveasfilename(
            initialdir=self.root_dir,
            defaultextension=".yaml",
            filetypes=[("YAML", "*.yaml"), ("YAML", "*.yml"), ("所有文件", "*.*")],
        )
        if not path:
            return
        self.template_var.set(self._display_path(Path(path)))
        self._save_template_from_editor()

    def _validate_template_editor(self) -> bool:
        ok, message = self._validate_template_text()
        if not ok:
            self.template_status_var.set(f"校验失败：{message}")
            messagebox.showerror("模板校验失败", message)
            return False
        self.template_status_var.set("模板校验通过")
        messagebox.showinfo("模板校验", "模板校验通过。")
        return True

    def _validate_template_text(self) -> tuple[bool, str]:
        from .workflow import validate_template

        try:
            template = yaml.safe_load(self.template_text.get("1.0", "end-1c")) or {}
            validate_template(template)
        except Exception as exc:
            return False, str(exc)
        return True, "模板校验通过"

    def _insert_template_snippet(self, name: str) -> None:
        snippet = ACTION_SNIPPETS[name]
        insert_at = self.template_text.index(tk.INSERT)
        before = self.template_text.get("1.0", insert_at)
        prefix = "" if before.endswith("\n") or not before else "\n"
        self.template_text.insert(insert_at, prefix + snippet)
        self.template_text.focus_set()
        self.template_status_var.set(f"已插入动作：{name}")

    def _preflight_check(self) -> None:
        checks: list[tuple[str, bool, str]] = []
        config_path = self._resolve_path(self.config_var.get().strip() or "config.yaml")
        template_path = self._resolve_path(self.template_var.get().strip() or "")
        checks.append(("配置文件", config_path.exists(), self._display_path(config_path)))
        checks.append(("流程模板", template_path.exists(), self._display_path(template_path)))

        ok, message = self._validate_template_text()
        checks.append(("模板结构", ok, message))

        data_source_message = "未解析"
        data_source_ok = False
        try:
            template = yaml.safe_load(self.template_text.get("1.0", "end-1c")) or {}
            data_file = Path(str(template.get("data_source", {}).get("file", "")))
            if data_file:
                if not data_file.is_absolute():
                    data_file = self.root_dir / data_file
                data_source_ok = data_file.exists()
                data_source_message = self._display_path(data_file)
        except Exception as exc:
            data_source_message = str(exc)
        checks.append(("数据源", data_source_ok, data_source_message))

        screen = f"{self.root.winfo_screenwidth()} x {self.root.winfo_screenheight()}"
        checks.append(("屏幕分辨率", True, screen))
        checks.append(("执行模式", True, "真实执行" if self.live_var.get() else "安全预演"))

        lines = ["运行前预检"]
        all_ok = True
        for name, passed, detail in checks:
            all_ok = all_ok and passed
            marker = "通过" if passed else "注意"
            lines.append(f"- {name}: {marker} ({detail})")
        report = "\n".join(lines)
        self._append_output(report + "\n")
        self.template_status_var.set("预检通过" if all_ok else "预检有注意项")
        if all_ok:
            messagebox.showinfo("运行前预检", report)
        else:
            messagebox.showwarning("运行前预检", report)

    def _base_command(self) -> list[str]:
        command = [sys.executable, str(self.main_path), "--config", self.config_var.get().strip() or "config.yaml"]
        if self.run_id_var.get().strip():
            command.extend(["--run-id", self.run_id_var.get().strip()])
        return command

    def _mode_args(self, force_dry_run: bool = False) -> list[str]:
        if force_dry_run or not self.live_var.get():
            return ["--dry-run"]
        return ["--live"]

    def _parse_region(self) -> list[str]:
        raw = self.region_var.get().strip()
        if not raw:
            return []
        parts = raw.replace(",", " ").split()
        if len(parts) != 4:
            raise ValueError("截图区域必须是四个数字：x1 y1 x2 y2。")
        return ["--region", *parts]

    def click_command(self, force_dry_run: bool = False) -> list[str]:
        text = self.click_text_var.get().strip()
        if not text:
            raise ValueError("请填写目标文字。")
        command = self._base_command()
        command.extend(["click", text, *self._mode_args(force_dry_run)])
        command.extend(["--match", MATCH_MODES.get(self.match_var.get(), "contains")])
        if self.confidence_var.get().strip():
            command.extend(["--min-confidence", self.confidence_var.get().strip()])
        command.extend(["--occurrence", self.occurrence_var.get().strip() or "1"])
        command.extend(self._parse_region())
        return command

    def workflow_command(self, force_dry_run: bool = False) -> list[str]:
        template = self.template_var.get().strip()
        if not template:
            raise ValueError("请选择流程模板。")
        if not self._save_template_from_editor(silent=True):
            raise ValueError("模板保存失败，请检查模板路径。")
        ok, message = self._validate_template_text()
        if not ok:
            raise ValueError(f"模板校验失败：{message}")
        command = self._base_command()
        command.extend(["run", template, *self._mode_args(force_dry_run)])
        if self.only_var.get().strip():
            command.extend(["--only", self.only_var.get().strip()])
        if self.result_var.get().strip():
            command.extend(["--result", self.result_var.get().strip()])
        return command

    def run_click(self) -> None:
        self._run_command(lambda: self.click_command(force_dry_run=False))

    def preview_click(self) -> None:
        self._run_command(lambda: self.click_command(force_dry_run=True))

    def run_workflow(self) -> None:
        self._run_command(lambda: self.workflow_command(force_dry_run=False))

    def preview_workflow(self) -> None:
        self._run_command(lambda: self.workflow_command(force_dry_run=True))

    def _run_command(self, builder: Callable[[], list[str]]) -> None:
        if self.process and self.process.poll() is None:
            self._append_output("已有命令正在运行。\n")
            return
        try:
            command = builder()
        except Exception as exc:
            self.status_var.set(str(exc))
            return
        self.command_preview.configure(text=self._format_command(command))
        self.output.delete("1.0", tk.END)
        self.status_var.set("运行中")
        self.stop_button.configure(state=tk.NORMAL)
        thread = threading.Thread(target=self._worker, args=(command,), daemon=True)
        thread.start()

    def _worker(self, command: list[str]) -> None:
        self.output_queue.put(f"$ {self._format_command(command)}\n")
        try:
            self.process = subprocess.Popen(
                command,
                cwd=self.root_dir,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                encoding="utf-8",
                errors="replace",
            )
            assert self.process.stdout is not None
            for line in self.process.stdout:
                self.output_queue.put(line)
            code = self.process.wait()
            self.output_queue.put(f"\n退出码：{code}\n")
            status = "就绪" if code == 0 else f"失败（退出码 {code}）"
            self.output_queue.put(f"__STATUS__:{status}\n")
        except Exception as exc:
            self.output_queue.put(f"\n{exc}\n")
            self.output_queue.put("__STATUS__:失败\n")
        finally:
            self.process = None

    def _poll_output(self) -> None:
        try:
            while True:
                text = self.output_queue.get_nowait()
                if text.startswith("__STATUS__:"):
                    self.status_var.set(text.removeprefix("__STATUS__:").strip())
                    self.stop_button.configure(state=tk.DISABLED)
                else:
                    self._append_output(text)
        except queue.Empty:
            pass
        self.root.after(80, self._poll_output)

    def _append_output(self, text: str) -> None:
        self.output.insert(tk.END, text)
        self.output.see(tk.END)

    def stop_process(self) -> None:
        if self.process and self.process.poll() is None:
            self.process.terminate()
            self.status_var.set("正在停止")

    def _format_command(self, command: list[str]) -> str:
        return " ".join(f'"{part}"' if " " in part else part for part in command)

    def _focus_click(self) -> None:
        self.click_card.focus_set()

    def _focus_workflow(self) -> None:
        self.workflow_card.focus_set()

    def _focus_template(self) -> None:
        self.template_text.focus_set()

    def _focus_log(self) -> None:
        self.output.focus_set()


def launch_app(default_config: str = "config.yaml") -> None:
    enable_dpi_awareness()
    root = tk.Tk()
    app = FluentRpaApp(root, default_config=default_config)
    root.mainloop()
