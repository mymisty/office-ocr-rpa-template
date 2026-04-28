from __future__ import annotations

import queue
import subprocess
import sys
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, ttk
from typing import Callable


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

        self.click_text_var = tk.StringVar(value="Submit")
        self.match_var = tk.StringVar(value="contains")
        self.confidence_var = tk.StringVar(value="")
        self.occurrence_var = tk.StringVar(value="1")
        self.region_var = tk.StringVar()

        self.template_var = tk.StringVar(value="tasks/batch_names.yaml")
        self.only_var = tk.StringVar()
        self.result_var = tk.StringVar()

        self._configure_window()
        self._configure_styles()
        self._build_layout()
        self._poll_output()

    def _configure_window(self) -> None:
        self.root.title("Office OCR RPA")
        self.root.geometry("1120x720")
        self.root.minsize(980, 640)
        self.root.configure(bg="#f3f3f3")

    def _configure_styles(self) -> None:
        self.style = ttk.Style(self.root)
        self.style.theme_use("clam")
        self.style.configure(".", font=("Segoe UI", 10), background="#f3f3f3", foreground="#1f1f1f")
        self.style.configure("App.TFrame", background="#f3f3f3")
        self.style.configure("Surface.TFrame", background="#ffffff", relief="flat")
        self.style.configure("Sidebar.TFrame", background="#fbfbfb")
        self.style.configure("Title.TLabel", font=("Segoe UI Variable Display", 22, "bold"), background="#f3f3f3")
        self.style.configure("Subtitle.TLabel", font=("Segoe UI", 10), background="#f3f3f3", foreground="#606060")
        self.style.configure("CardTitle.TLabel", font=("Segoe UI", 12, "bold"), background="#ffffff")
        self.style.configure("CardText.TLabel", background="#ffffff", foreground="#4d4d4d")
        self.style.configure("TLabel", background="#ffffff")
        self.style.configure("TCheckbutton", background="#ffffff")
        self.style.configure("TEntry", fieldbackground="#ffffff", bordercolor="#d0d0d0", lightcolor="#d0d0d0")
        self.style.configure("TCombobox", fieldbackground="#ffffff", bordercolor="#d0d0d0")
        self.style.configure("Accent.TButton", background="#0078d4", foreground="#ffffff", bordercolor="#0078d4")
        self.style.map("Accent.TButton", background=[("active", "#106ebe"), ("disabled", "#c8c8c8")])
        self.style.configure("Ghost.TButton", background="#ffffff", bordercolor="#d0d0d0")
        self.style.map("Ghost.TButton", background=[("active", "#f3f3f3")])
        self.style.configure("Nav.TButton", anchor="w", background="#fbfbfb", bordercolor="#fbfbfb", padding=(16, 10))
        self.style.map("Nav.TButton", background=[("active", "#eef6fc")])

    def _build_layout(self) -> None:
        shell = ttk.Frame(self.root, style="App.TFrame", padding=16)
        shell.pack(fill=tk.BOTH, expand=True)
        shell.columnconfigure(1, weight=1)
        shell.rowconfigure(0, weight=1)

        sidebar = ttk.Frame(shell, style="Sidebar.TFrame", padding=(10, 14))
        sidebar.grid(row=0, column=0, sticky="nsew", padx=(0, 14))
        sidebar.configure(width=220)
        sidebar.grid_propagate(False)

        brand = ttk.Label(sidebar, text="Office OCR", font=("Segoe UI Variable Display", 18, "bold"), background="#fbfbfb")
        brand.pack(anchor="w", padx=8, pady=(0, 22))
        ttk.Button(sidebar, text="Quick click", style="Nav.TButton", command=self._focus_click).pack(fill=tk.X, pady=2)
        ttk.Button(sidebar, text="Workflow", style="Nav.TButton", command=self._focus_workflow).pack(fill=tk.X, pady=2)
        ttk.Button(sidebar, text="Run log", style="Nav.TButton", command=self._focus_log).pack(fill=tk.X, pady=2)
        ttk.Frame(sidebar, style="Sidebar.TFrame").pack(fill=tk.BOTH, expand=True)
        ttk.Label(sidebar, textvariable=self._mode_label(), background="#fbfbfb", foreground="#606060").pack(anchor="w", padx=8)

        content = ttk.Frame(shell, style="App.TFrame")
        content.grid(row=0, column=1, sticky="nsew")
        content.columnconfigure(0, weight=1)
        content.rowconfigure(2, weight=1)

        header = ttk.Frame(content, style="App.TFrame")
        header.grid(row=0, column=0, sticky="ew", pady=(0, 14))
        header.columnconfigure(0, weight=1)
        ttk.Label(header, text="Automation Console", style="Title.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(header, text="Screen OCR, workflow runs, and execution output", style="Subtitle.TLabel").grid(
            row=1,
            column=0,
            sticky="w",
            pady=(2, 0),
        )

        self.global_card = self._card(content, row=1, title="Run settings")
        self.global_card.columnconfigure(1, weight=1)
        self._field(self.global_card, "Config", self.config_var, 0, browse=lambda: self._browse_file(self.config_var, "YAML", "*.yaml"))
        self._field(self.global_card, "Run id", self.run_id_var, 1)
        ttk.Checkbutton(self.global_card, text="Live actions", variable=self.live_var).grid(row=2, column=1, sticky="w", pady=(8, 0))

        work_area = ttk.Frame(content, style="App.TFrame")
        work_area.grid(row=2, column=0, sticky="nsew")
        work_area.columnconfigure(0, weight=1)
        work_area.columnconfigure(1, weight=1)
        work_area.rowconfigure(1, weight=1)

        self.click_card = self._card(work_area, row=0, column=0, title="Quick click", padx=(0, 7))
        self._field(self.click_card, "Target text", self.click_text_var, 0)
        self._combo(self.click_card, "Match", self.match_var, ("contains", "exact", "fuzzy"), 1)
        self._field(self.click_card, "Min confidence", self.confidence_var, 2)
        self._field(self.click_card, "Occurrence", self.occurrence_var, 3)
        self._field(self.click_card, "Region", self.region_var, 4)
        self._button_row(self.click_card, 5, ("Run click", self.run_click, "Accent.TButton"), ("Preview", self.preview_click, "Ghost.TButton"))

        self.workflow_card = self._card(work_area, row=0, column=1, title="Workflow", padx=(7, 0))
        self._field(
            self.workflow_card,
            "Template",
            self.template_var,
            0,
            browse=lambda: self._browse_file(self.template_var, "YAML", "*.yaml"),
        )
        self._field(self.workflow_card, "Only status", self.only_var, 1)
        self._field(
            self.workflow_card,
            "Result file",
            self.result_var,
            2,
            browse=lambda: self._browse_save_file(self.result_var),
        )
        self._button_row(
            self.workflow_card,
            3,
            ("Run workflow", self.run_workflow, "Accent.TButton"),
            ("Dry run", self.preview_workflow, "Ghost.TButton"),
        )

        log_card = self._card(work_area, row=1, column=0, columnspan=2, title="Run log", pady=(14, 0))
        log_card.rowconfigure(1, weight=1)
        log_card.columnconfigure(0, weight=1)
        self.command_preview = ttk.Label(log_card, text="", style="CardText.TLabel")
        self.command_preview.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        self.output = tk.Text(
            log_card,
            height=10,
            wrap="word",
            relief="flat",
            bg="#1e1e1e",
            fg="#f5f5f5",
            insertbackground="#ffffff",
            font=("Cascadia Mono", 10),
        )
        self.output.grid(row=1, column=0, sticky="nsew")
        scrollbar = ttk.Scrollbar(log_card, command=self.output.yview)
        scrollbar.grid(row=1, column=1, sticky="ns")
        self.output.configure(yscrollcommand=scrollbar.set)

        footer = ttk.Frame(content, style="App.TFrame")
        footer.grid(row=3, column=0, sticky="ew", pady=(12, 0))
        footer.columnconfigure(0, weight=1)
        self.status_var = tk.StringVar(value="Ready")
        ttk.Label(footer, textvariable=self.status_var, style="Subtitle.TLabel").grid(row=0, column=0, sticky="w")
        self.stop_button = ttk.Button(footer, text="Stop", style="Ghost.TButton", command=self.stop_process, state=tk.DISABLED)
        self.stop_button.grid(row=0, column=1, sticky="e")

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
            ttk.Button(parent, text="Browse", style="Ghost.TButton", command=browse).grid(row=grid_row, column=2, padx=(8, 0), pady=6)

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
        mode = tk.StringVar(value="Default: dry run")

        def refresh(*_: object) -> None:
            mode.set("Mode: live actions" if self.live_var.get() else "Default: dry run")

        self.live_var.trace_add("write", refresh)
        return mode

    def _browse_file(self, variable: tk.StringVar, label: str, pattern: str) -> None:
        path = filedialog.askopenfilename(initialdir=self.root_dir, filetypes=[(label, pattern), ("All files", "*.*")])
        if path:
            variable.set(self._display_path(Path(path)))

    def _browse_save_file(self, variable: tk.StringVar) -> None:
        path = filedialog.asksaveasfilename(
            initialdir=self.root_dir,
            defaultextension=".xlsx",
            filetypes=[("Excel workbook", "*.xlsx"), ("All files", "*.*")],
        )
        if path:
            variable.set(self._display_path(Path(path)))

    def _display_path(self, path: Path) -> str:
        try:
            return str(path.resolve().relative_to(self.root_dir))
        except ValueError:
            return str(path)

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
            raise ValueError("Region must be four numbers: x1 y1 x2 y2.")
        return ["--region", *parts]

    def click_command(self, force_dry_run: bool = False) -> list[str]:
        text = self.click_text_var.get().strip()
        if not text:
            raise ValueError("Target text is required.")
        command = self._base_command()
        command.extend(["click", text, *self._mode_args(force_dry_run)])
        command.extend(["--match", self.match_var.get()])
        if self.confidence_var.get().strip():
            command.extend(["--min-confidence", self.confidence_var.get().strip()])
        command.extend(["--occurrence", self.occurrence_var.get().strip() or "1"])
        command.extend(self._parse_region())
        return command

    def workflow_command(self, force_dry_run: bool = False) -> list[str]:
        template = self.template_var.get().strip()
        if not template:
            raise ValueError("Template is required.")
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
            self._append_output("A command is already running.\n")
            return
        try:
            command = builder()
        except Exception as exc:
            self.status_var.set(str(exc))
            return
        self.command_preview.configure(text=self._format_command(command))
        self.output.delete("1.0", tk.END)
        self.status_var.set("Running")
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
            self.output_queue.put(f"\nExit code: {code}\n")
            self.output_queue.put("__STATUS__:Ready\n")
        except Exception as exc:
            self.output_queue.put(f"\n{exc}\n")
            self.output_queue.put("__STATUS__:Failed\n")
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
            self.status_var.set("Stopping")

    def _format_command(self, command: list[str]) -> str:
        return " ".join(f'"{part}"' if " " in part else part for part in command)

    def _focus_click(self) -> None:
        self.click_card.focus_set()

    def _focus_workflow(self) -> None:
        self.workflow_card.focus_set()

    def _focus_log(self) -> None:
        self.output.focus_set()


def launch_app(default_config: str = "config.yaml") -> None:
    root = tk.Tk()
    app = FluentRpaApp(root, default_config=default_config)
    root.mainloop()
