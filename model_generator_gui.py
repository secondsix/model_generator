"""物模型导入文件生成器的桌面界面。"""

import os
import json
import queue
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from tag_generator import generate_tag_files

from model_generator import (
    DEFAULT_OUTPUT_DIR,
    generate_many_files,
)


class ModelGeneratorApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("物模型标签库导入生成器")
        self.geometry("820x540")
        self.minsize(720, 480)

        self.source_var = tk.StringVar()
        self.template_var = tk.StringVar()
        self.output_var = tk.StringVar()
        self.status_var = tk.StringVar(value="请选择文件和导出位置")
        self.progress_var = tk.DoubleVar(value=0)
        self.events = queue.Queue()
        self.mode_var = tk.StringVar(value="物模型标签库导入生成器")
        self.mode_paths = {}
        self.current_mode = self.mode_var.get()

        self._build_ui()
        self._switch_mode()
        self.after(100, self._process_events)

    def _build_ui(self):
        self.columnconfigure(0, weight=1)
        self.rowconfigure(1, weight=1)

        self.mode_selector = ttk.Combobox(self, textvariable=self.mode_var, state="readonly",
            values=("物模型标签库导入生成器", "物模型导入文件生成器"), width=34,
            font=("Microsoft YaHei UI", 16, "bold"))
        self.mode_selector.grid(row=0, column=0, padx=28, pady=(24, 12), sticky="w")
        self.mode_selector.bind("<<ComboboxSelected>>", self._switch_mode)

        body = ttk.Frame(self, padding=(28, 10, 28, 24))
        body.grid(row=1, column=0, sticky="nsew")
        body.columnconfigure(1, weight=1)
        body.rowconfigure(5, weight=1)

        self.source_button = self._path_row(body, 0, "标签库文件", self.source_var, self._choose_source, button_text="选择文件（多选）")
        self._path_row(body, 1, "导入模板", self.template_var, self._choose_template)
        self.output_button = self._path_row(body, 2, "导出位置", self.output_var, self._choose_output, button_text="选择文件夹")

        ttk.Separator(body).grid(row=3, column=0, columnspan=3, sticky="ew", pady=(20, 16))

        action_bar = ttk.Frame(body)
        action_bar.grid(row=4, column=0, columnspan=3, sticky="ew")
        action_bar.columnconfigure(1, weight=1)
        self.generate_button = ttk.Button(action_bar, text="开始生成", command=self._start_generation)
        self.generate_button.grid(row=0, column=0, sticky="w")
        ttk.Button(action_bar, text="打开导出位置", command=self._open_output).grid(row=0, column=2, sticky="e")

        log_frame = ttk.LabelFrame(body, text="处理结果", padding=10)
        log_frame.grid(row=5, column=0, columnspan=3, sticky="nsew", pady=(16, 12))
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)
        self.log = tk.Text(log_frame, height=12, wrap="word", state="disabled", font=("Microsoft YaHei UI", 10))
        self.log.grid(row=0, column=0, sticky="nsew")
        scrollbar = ttk.Scrollbar(log_frame, orient="vertical", command=self.log.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        self.log.configure(yscrollcommand=scrollbar.set)

        self.progress = ttk.Progressbar(body, variable=self.progress_var, maximum=100)
        self.progress.grid(row=6, column=0, columnspan=3, sticky="ew")
        ttk.Label(body, textvariable=self.status_var).grid(row=7, column=0, columnspan=3, sticky="w", pady=(8, 0))

    def _path_row(self, parent, row, label, variable, command, button_text="选择文件"):
        label_widget = ttk.Label(parent, text=label, width=10)
        label_widget.grid(row=row, column=0, sticky="w", pady=7)
        if row == 0:
            self.source_label = label_widget
        ttk.Entry(parent, textvariable=variable).grid(row=row, column=1, sticky="ew", padx=(0, 10), pady=7)
        button = ttk.Button(parent, text=button_text, command=command)
        button.grid(row=row, column=2, sticky="e", pady=7)
        return button

    def _switch_mode(self, event=None):
        self.mode_paths[self.current_mode] = (self.source_var.get(), self.template_var.get(), self.output_var.get())
        self.current_mode = self.mode_var.get()
        source, template, output = self.mode_paths.get(self.current_mode, ("", "", ""))
        self.source_var.set(source)
        self.source_label.configure(text="设计表" if self._is_tag_mode() else "标签库文件")
        self.source_button.configure(text="选择文件" if self._is_tag_mode() else "选择文件（多选）")
        self.template_var.set(template)
        self.output_var.set(output)
        self.title(self.current_mode)
        self.output_button.configure(text="选择文件夹")
        self.progress_var.set(0)
        self.status_var.set("请选择文件和导出位置")
        self._set_log("标签类型：property；标识：GY_分类标识_名称拼音首字母（大写）\n分类默认首字母，冲突时用完整拼音；按产品分别导出；仅处理参数类型；同名标签追加单位后缀，数字首字母冲突时展开全拼；跨产品重名追加 _产品名称；说明取备注说明。\n" if self._is_tag_mode() else "")

    def _is_tag_mode(self):
        return self.mode_var.get() == "物模型标签库导入生成器"

    def _choose_source(self):
        current = self.source_var.get().strip()
        if not self._is_tag_mode():
            try:
                paths = self._source_paths()
                current = str(paths[0]) if paths else ""
            except (ValueError, TypeError):
                current = ""
        chooser = filedialog.askopenfilename if self._is_tag_mode() else filedialog.askopenfilenames
        selected = chooser(
            title="选择物模型设计表" if self._is_tag_mode() else "选择物模型标签库导入文件",
            initialdir=str(Path(current).parent) if current else None,
            filetypes=[("Excel 工作簿", "*.xlsx"), ("所有文件", "*.*")],
        )
        if selected:
            self.source_var.set(selected if self._is_tag_mode() else json.dumps(list(selected), ensure_ascii=False))
            if not self._is_tag_mode():
                self.status_var.set(f"已选择 {len(selected)} 个标签库文件")

    def _source_paths(self):
        text = self.source_var.get().strip()
        if not text:
            return []
        values = json.loads(text) if text.startswith("[") else [text]
        if not isinstance(values, list) or not all(isinstance(path, str) and path.strip() for path in values):
            raise ValueError("请选择有效的标签库文件。")
        return [Path(path) for path in values]

    def _choose_template(self):
        current = self.template_var.get().strip()
        selected = filedialog.askopenfilename(
            title="选择物模型标签库模板" if self._is_tag_mode() else "选择物模型导入模板",
            initialdir=str(Path(current).parent) if current else None,
            filetypes=[("Excel 工作簿", "*.xlsx"), ("所有文件", "*.*")],
        )
        if selected:
            self.template_var.set(selected)

    def _choose_output(self):
        selected = filedialog.askdirectory(
            title="选择导出文件的存储位置",
            initialdir=self.output_var.get(),
        )
        if selected:
            self.output_var.set(selected)

    def _start_generation(self):
        source_text = self.source_var.get().strip()
        template_text = self.template_var.get().strip()
        output_text = self.output_var.get().strip()
        if not output_text:
            messagebox.showerror("无法开始", "请选择导出文件的存储位置。")
            return
        try:
            source = Path(source_text) if self._is_tag_mode() else self._source_paths()
        except (ValueError, TypeError) as exc:
            messagebox.showerror("无法开始", str(exc))
            return
        template = Path(template_text)
        output = Path(output_text)
        sources = [source] if self._is_tag_mode() else source
        if not sources or any(not path.is_file() for path in sources):
            messagebox.showerror("无法开始", "请选择有效的源 Excel 文件。")
            return
        if not template.is_file():
            messagebox.showerror("无法开始", "请选择有效的导入模板。")
            return
        tag_mode = self._is_tag_mode()
        if output.exists() and not output.is_dir():
            messagebox.showerror("无法开始", "请选择文件夹作为导出位置。")
            return
        if tag_mode and output.is_dir() and any(output.glob("*物模型标签库导入.xlsx")):
            if not messagebox.askyesno("覆盖文件", "导出目录中已有标签库文件，是否覆盖本次生成的同名文件？"):
                return
        self.generate_button.configure(state="disabled")
        self.mode_selector.configure(state="disabled")
        self.progress_var.set(0)
        self.status_var.set("正在读取源文件……")
        self._set_log("")
        threading.Thread(
            target=self._run_generation,
            args=(source, template, output, tag_mode),
            daemon=True,
        ).start()

    def _run_generation(self, source, template, output, tag_mode=False):
        try:
            def progress(index, total, model_name, output_file):
                self.events.put(("progress", index, total, model_name, output_file))

            generator = generate_tag_files if tag_mode else generate_many_files
            result = generator(source, template, output, progress=progress)
            self.events.put(("done", result))
        except Exception as exc:
            self.events.put(("error", str(exc)))

    def _process_events(self):
        try:
            while True:
                event = self.events.get_nowait()
                if event[0] == "progress":
                    _, index, total, model_name, output_file = event
                    self.progress_var.set(index / total * 100 if total else 0)
                    self.status_var.set(f"正在生成：{index}/{total}  {model_name}")
                    self._append_log(f"[{index}/{total}] {output_file.name}\n")
                elif event[0] == "done":
                    result = event[1]
                    count = len(result["generated_files"])
                    self.progress_var.set(100)
                    self.status_var.set(f"生成完成，共 {count} 个文件")
                    self.generate_button.configure(state="normal")
                    self.mode_selector.configure(state="readonly")
                    if "tag_count" in result:
                        self._append_log(f"共生成 {result['tag_count']} 条标签。\n")
                    messagebox.showinfo("生成完成", f"已生成 {count} 个文件。\n\n存储位置：\n{result['output_dir']}")
                elif event[0] == "error":
                    self.generate_button.configure(state="normal")
                    self.mode_selector.configure(state="readonly")
                    self.status_var.set("生成失败")
                    self._append_log(f"错误：{event[1]}\n")
                    messagebox.showerror("生成失败", event[1])
        except queue.Empty:
            pass
        self.after(100, self._process_events)

    def _open_output(self):
        output = Path(self.output_var.get().strip())
        if not output.exists():
            messagebox.showwarning("无法打开", "导出位置尚不存在。")
            return
        os.startfile(output)

    def _set_log(self, value):
        self.log.configure(state="normal")
        self.log.delete("1.0", "end")
        self.log.insert("end", value)
        self.log.configure(state="disabled")

    def _append_log(self, value):
        self.log.configure(state="normal")
        self.log.insert("end", value)
        self.log.see("end")
        self.log.configure(state="disabled")


if __name__ == "__main__":
    ModelGeneratorApp().mainloop()
