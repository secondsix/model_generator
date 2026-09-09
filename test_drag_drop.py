import tempfile
import unittest
from pathlib import Path
from model_generator_gui import ModelGeneratorApp
from tkinterdnd2 import COPY, REFUSE_DROP


class DragDropTests(unittest.TestCase):
    def test_files_modes_and_validation(self):
        app = ModelGeneratorApp()
        app.withdraw()
        try:
            self.assertTrue(app.tk.call("package", "present", "tkdnd"))
            with tempfile.TemporaryDirectory() as directory:
                root = Path(directory)
                files = [root / "中文 文件.xlsx", root / "第二个.xlsx"]
                for path in files: path.touch()
                single = "{" + str(files[0]) + "}"
                multiple = " ".join("{" + str(path) + "}" for path in files)
                self.assertEqual(app._drop_paths(single, app.source_var), COPY)
                self.assertEqual(app.source_var.get(), str(files[0]))
                self.assertEqual(app._drop_paths(multiple, app.source_var), REFUSE_DROP)
                self.assertEqual(app.source_var.get(), str(files[0]))
                for variable in (app.template_var, app.device_var):
                    self.assertEqual(app._drop_paths(single, variable), COPY)
                self.assertEqual(app._drop_paths("{" + str(root) + "}", app.output_var), COPY)
                self.assertEqual(app._drop_paths(single, app.output_var), REFUSE_DROP)
                app.mode_var.set("物模型导入文件生成器")
                app._switch_mode()
                self.assertEqual(app._drop_paths(multiple, app.source_var), COPY)
                self.assertEqual(app._source_paths(), files)
                app.generate_button.configure(state="disabled")
                self.assertEqual(app._drop_paths(single, app.source_var), REFUSE_DROP)
                self.assertEqual(app._source_paths(), files)
        finally:
            app.destroy()
