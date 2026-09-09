"""物模型导入文件生成器的桌面界面。"""

import os
import json
import queue
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from tkinterdnd2 import TkinterDnD, DND_FILES, COPY, REFUSE_DROP
from tag_generator import generate_tag_files
from point_generator import generate_point_file

from model_generator import (
    DEFAULT_OUTPUT_DIR,
    generate_many_files,
)


class ModelGeneratorApp(TkinterDnD.Tk):
    def __init__(self):
        super().__init__()
        self.title("物模型标签库导入生成器")
        self.geometry("900x600")
        self.minsize(820, 540)

        self.source_var = tk.StringVar()
        self.template_var = tk.StringVar()
        self.device_var = tk.StringVar()
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
            values=("物模型标签库导入生成器", "物模型导入文件生成器", "单色注塑机点位导入生成器"), width=34,
            font=("Microsoft YaHei UI", 16, "bold"))
        self.mode_selector.grid(row=0, column=0, padx=28, pady=(24, 12), sticky="w")
        self.mode_selector.bind("<<ComboboxSelected>>", self._switch_mode)

        body = ttk.Frame(self, padding=(28, 10, 28, 24))
        body.grid(row=1, column=0, sticky="nsew")
        body.columnconfigure(1, weight=1)
        body.rowconfigure(6, weight=1)

        self.source_button = self._path_row(body, 0, "标签库文件", self.source_var, self._choose_source, button_text="选择文件（多选）")
        self._path_row(body, 1, "导入模板", self.template_var, self._choose_template)
        self.output_button = self._path_row(body, 3, "导出位置", self.output_var, self._choose_output, button_text="选择文件夹")

        self.device_frame = ttk.Frame(body)
        self.device_frame.grid(row=2, column=0, columnspan=3, sticky="ew")
        self.device_frame.columnconfigure(1, weight=1)
        ttk.Label(self.device_frame, text="设备点表", width=10).grid(row=0, column=0, sticky="w")
        device_entry = ttk.Entry(self.device_frame, textvariable=self.device_var)
        device_entry.grid(row=0, column=1, sticky="ew", padx=(0, 10), pady=7)
        device_button = ttk.Button(self.device_frame, text="选择文件", command=self._choose_device)
        device_button.grid(row=0, column=2)
        self._register_drop(device_entry, self.device_var)
        self._register_drop(device_button, self.device_var)

        ttk.Separator(body).grid(row=4, column=0, columnspan=3, sticky="ew", pady=(20, 16))

        action_bar = ttk.Frame(body)
        action_bar.grid(row=5, column=0, columnspan=3, sticky="ew")
        action_bar.columnconfigure(1, weight=1)
        self.generate_button = ttk.Button(action_bar, text="开始生成", command=self._start_generation)
        self.generate_button.grid(row=0, column=0, sticky="w")
        ttk.Button(action_bar, text="打开导出位置", command=self._open_output).grid(row=0, column=2, sticky="e")

        log_frame = ttk.LabelFrame(body, text="处理结果", padding=10)
        log_frame.grid(row=6, column=0, columnspan=3, sticky="nsew", pady=(16, 12))
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)
        self.log = tk.Text(log_frame, height=12, wrap="word", state="disabled", font=("Microsoft YaHei UI", 10))
        self.log.grid(row=0, column=0, sticky="nsew")
        scrollbar = ttk.Scrollbar(log_frame, orient="vertical", command=self.log.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        self.log.configure(yscrollcommand=scrollbar.set)

        self.progress = ttk.Progressbar(body, variable=self.progress_var, maximum=100)
        self.progress.grid(row=7, column=0, columnspan=3, sticky="ew")
        ttk.Label(body, textvariable=self.status_var).grid(row=8, column=0, columnspan=3, sticky="w", pady=(8, 0))

    def _path_row(self, parent, row, label, variable, command, button_text="选择文件"):
        label_widget = ttk.Label(parent, text=label, width=10)
        label_widget.grid(row=row, column=0, sticky="w", pady=7)
        if row == 0:
            self.source_label = label_widget
        entry = ttk.Entry(parent, textvariable=variable)
        entry.grid(row=row, column=1, sticky="ew", padx=(0, 10), pady=7)
        button = ttk.Button(parent, text=button_text, command=command)
        button.grid(row=row, column=2, sticky="e", pady=7)
        self._register_drop(entry, variable)
        self._register_drop(button, variable)
        return button

    def _register_drop(self, widget, variable):
        widget.drop_target_register(DND_FILES)
        widget.dnd_bind("<<Drop>>", lambda event: self._drop_paths(event.data, variable))

    def _drop_paths(self, data, variable):
        if str(self.generate_button.cget("state")) == "disabled":
            self.status_var.set("正在生成，请完成后再拖入文件")
            return REFUSE_DROP
        try:
            paths = list(dict.fromkeys(self.tk.splitlist(data)))
            if not paths:
                return REFUSE_DROP
            multiple = variable is self.source_var and not self._single_source()
            if len(paths) > 1 and not multiple:
                raise ValueError("此处只能拖入一个文件或文件夹。")
            if variable is self.output_var:
                if not Path(paths[0]).is_dir():
                    raise ValueError("请向导出位置拖入文件夹。")
            elif any(not Path(path).is_file() or Path(path).suffix.lower() != ".xlsx" for path in paths):
                raise ValueError("请拖入有效的 .xlsx 文件。")
            variable.set(json.dumps(paths, ensure_ascii=False) if multiple else paths[0])
            self.status_var.set(f"已拖入 {len(paths)} 个标签库文件" if multiple else f"已选择：{Path(paths[0]).name}")
            return COPY
        except (ValueError, tk.TclError) as exc:
            self.status_var.set(str(exc))
            return REFUSE_DROP

    def _switch_mode(self, event=None):
        self.mode_paths[self.current_mode] = (self.source_var.get(), self.template_var.get(), self.output_var.get(), self.device_var.get())
        self.current_mode = self.mode_var.get()
        source, template, output, device = self.mode_paths.get(self.current_mode, ("", "", "", ""))
        self.device_var.set(device)
        if self._is_point_mode():
            self.device_frame.grid()
        else:
            self.device_frame.grid_remove()
        self.source_var.set(source)
        self.source_label.configure(text="工艺参数" if self._is_point_mode() else ("设计表" if self._is_tag_mode() else "标签库文件"))
        self.source_button.configure(text="选择文件" if self._single_source() else "选择文件（多选）")
        self.template_var.set(template)
        self.output_var.set(output)
        self.title(self.current_mode)
        self.output_button.configure(text="选择文件夹")
        self.progress_var.set(0)
        self.status_var.set("可将文件拖入对应输入框或选择按钮；导出位置可拖入文件夹")
        self._set_log("标签类型：property；标识：GY_分类标识_名称拼音首字母（大写）\n分类默认首字母，冲突时用完整拼音；按产品分别导出；仅处理参数类型；同名标签追加单位后缀，数字首字母冲突时展开全拼；跨产品重名追加 _产品名称；说明取备注说明。\n" if self._is_tag_mode() else "")

        if self._is_point_mode():
            self._set_log("按工艺参数文件名自动选择单色/双色规则；未匹配的地址和类型留空；采集周期 1000 ms。\n")

    def _is_tag_mode(self):
        return self.mode_var.get() == "物模型标签库导入生成器"

    def _is_point_mode(self):
        return self.mode_var.get() == "单色注塑机点位导入生成器"

    def _single_source(self):
        return self._is_tag_mode() or self._is_point_mode()

    def _choose_device(self):
        selected = filedialog.askopenfilename(title="选择 WinCC 设备点表", filetypes=[("Excel 工作簿", "*.xlsx")])
        if selected:
            self.device_var.set(selected)

    def _choose_source(self):
        current = self.source_var.get().strip()
        if not self._single_source():
            try:
                paths = self._source_paths()
                current = str(paths[0]) if paths else ""
            except (ValueError, TypeError):
                current = ""
        chooser = filedialog.askopenfilename if self._single_source() else filedialog.askopenfilenames
        selected = chooser(
            title="选择单色或双色注塑机工艺参数" if self._is_point_mode() else ("选择物模型设计表" if self._is_tag_mode() else "选择物模型标签库导入文件"),
            initialdir=str(Path(current).parent) if current else None,
            filetypes=[("Excel 工作簿", "*.xlsx"), ("所有文件", "*.*")],
        )
        if selected:
            self.source_var.set(selected if self._single_source() else json.dumps(list(selected), ensure_ascii=False))
            if not self._single_source():
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
            title="选择点位导入模板" if self._is_point_mode() else ("选择物模型标签库模板" if self._is_tag_mode() else "选择物模型导入模板"),
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
            source = Path(source_text) if self._single_source() else self._source_paths()
        except (ValueError, TypeError) as exc:
            messagebox.showerror("无法开始", str(exc))
            return
        template = Path(template_text)
        output = Path(output_text)
        sources = [source] if self._single_source() else source
        if not sources or any(not path.is_file() for path in sources):
            messagebox.showerror("无法开始", "请选择有效的源 Excel 文件。")
            return
        if not template.is_file():
            messagebox.showerror("无法开始", "请选择有效的导入模板。")
            return
        tag_mode = self._is_tag_mode()
        device = Path(self.device_var.get().strip()) if self._is_point_mode() else None
        if device is not None and not device.is_file():
            messagebox.showerror("无法开始", "请选择有效的设备点表。")
            return
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
            args=(source, template, output, tag_mode, device),
            daemon=True,
        ).start()

    def _run_generation(self, source, template, output, tag_mode=False, device=None):
        try:
            def progress(index, total, model_name, output_file):
                self.events.put(("progress", index, total, model_name, output_file))

            generator = generate_tag_files if tag_mode else generate_many_files
            result = generate_point_file(source, device, template, output, progress=progress) if device is not None else generator(source, template, output, progress=progress)
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
                    if "point_count" in result:
                        self._append_log(f"{result['rule']}规则：共 {result['point_count']} 个点位，完整匹配 {result['matched_count']} 个。\n")
                        for warning in result["warnings"]:
                            self._append_log(warning + "\n")
                        self.status_var.set(f"生成完成，匹配 {result['matched_count']}/{result['point_count']}；留空项见处理结果")
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
