生成定影模块内加热控制表

先导入依赖：pip install -r requirements.txt
再编译：python -m nuitka --mingw64 --standalone --include-package=openpyxl --output-dir=dist main.py
单文件版：python -m nuitka --mingw64 --standalone --onefile --include-package=openpyxl --output-dir=dist main.py