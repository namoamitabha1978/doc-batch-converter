# 文档批量转换工具

支持将 Word / Excel / PowerPoint 文档批量转换为 **PDF / PNG 序列 / JPG 序列** 的 GUI 工具。

## 功能特性

- **多格式输入**：`.docx`（Word）、`.xlsx`（Excel）、`.pptx`（PowerPoint）
- **批量处理**：支持选择单个文件或整个文件夹（递归扫描）
- **三种输出格式**：
  - `PDF` — 直接输出 PDF 文件
  - `PNG 序列` — 每页导出为 PNG 图片
  - `JPG 序列` — 每页导出为 JPG 图片（质量 90%）
- **可调 DPI**：72 ~ 300（默认 150）
- **自定义前缀**：输出文件名可加前缀
- **实时进度**：进度条 + 日志窗口
- **中途取消**：转换过程中可随时终止

## 依赖

- **LibreOffice**（用于文档 → PDF 转换）
  - 下载：https://www.libreoffice.org/download/
  - 默认路径：`C:\Program Files\LibreOffice\program\soffice.exe`
- **Python 3.10+**（源码运行时需要）
  - `pip install pymupdf`

## 快速开始

### 直接运行 EXE（推荐）

双击 `批量转换工具.exe` 即可使用，无需安装 Python。

### 从源码运行

```bash
pip install pymupdf
python batch_converter_gui.py
```

## 使用说明

1. **输入**：点击「选择文件…」或「选择文件夹…」
2. **输出目录**：选择转换结果保存位置
3. **输出格式**：PDF / PNG 序列 / JPG 序列
4. **DPI**：控制图片清晰度，150 适合屏幕查看，300 适合打印
5. 点击「开始转换」

## 打包为 EXE

```bash
pip install pymupdf pyinstaller

pyinstaller --onefile --windowed ^
  --name "批量转换工具" ^
  --exclude-module onnxruntime ^
  --exclude-module pymupdf_layout ^
  --exclude-module pymupdf4llm ^
  --exclude-module numpy ^
  --collect-submodules fitz ^
  batch_converter_gui.py
```

打包后 exe 在 `dist/` 目录下，约 30 MB。

## 工作原理

```
docx / xlsx / pptx
        ↓  (LibreOffice headless)
       PDF
        ↓  (PyMuPDF 渲染)
  PNG / JPG 序列
```

## 项目结构

```
doc-batch-converter/
├── batch_converter_gui.py   # 主程序（GUI + 转换逻辑）
├── README.md               # 本文件
└── .gitignore
```

## License

MIT
