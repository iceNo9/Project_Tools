生成定影模块内热敏电阻温度-电压映射关系的代码

先导入依赖：pip install -r requirements.txt
再编译：python -m nuitka --mingw64 --standalone --include-package=openpyxl --output-dir=dist main.py
单文件版：python -m nuitka --mingw64 --standalone --onefile --include-package=openpyxl --output-dir=dist main.py