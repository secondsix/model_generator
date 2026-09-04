推荐使用 PyInstaller。当前虚拟环境尚未安装它。在项目目录打开 PowerShell，依次执行：

```powershell
.\.venv\Scripts\python.exe -m pip install pyinstaller
```

然后打包：

```powershell
.\.venv\Scripts\python.exe -m PyInstaller `
  --noconfirm --clean --onefile --windowed `
  --name "物模型导入文件生成器" `
  model_generator_gui.py
```

调试时建议先去掉 `--windowed`，这样启动错误会显示在命令行窗口中：

```powershell
.\.venv\Scripts\python.exe -m PyInstaller `
  --noconfirm --clean --onefile `
  --name "物模型导入文件生成器-调试版" `
  model_generator_gui.py
```

