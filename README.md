# 批量转换工具（文档 ↔ 图片）

> 南无阿弥陀佛 🙏

一个带 GUI 的批量格式转换工具，支持 **文档转图片** 和 **图片合并为PDF** 双向操作。

---

## 功能总览

| 方向 | 输入格式 | 输出格式 | 说明 |
|------|---------|---------|------|
| 文档 → PDF | docx / xlsx / pptx | PDF | 批量转换为 PDF 文件 |
| 文档 → PNG 序列 | docx / xlsx / pptx | PNG 图片（每页一张） | 高分辨率逐页导出 |
| 文档 → JPG 序列 | docx / xlsx / pptx | JPG 图片（每页一张） | 体积更小，质量90% |
| **图片 → 合并PDF** | png / jpg / bmp / tiff | 单个 PDF | 按文件名排序合并 |

---

## 快速开始

### 方式一：直接运行 EXE（推荐）

双击 `批量转换工具.exe` 即可使用，无需安装 Python 环境。

> ⚠️ 运行文档转换功能需要本机已安装 **LibreOffice**。
> 下载地址：https://www.libreoffice.org/download/

### 方式二：从源码运行

```bash
# 安装依赖
pip install pymupdf

# 运行
python batch_converter_gui.py
```

---

## 使用说明

### 基本操作步骤

1. **选择转换模式**：在「转换模式」下拉框中选择
   - `文档 → PDF`
   - `文档 → PNG 序列`
   - `文档 → JPG 序列`
   - `图片 → 合并为 PDF`

2. **选择输入**：点击「选择文件…」或「选择文件夹…」
   - 文件夹模式会**递归扫描**所有子目录中的匹配文件

3. **选择输出目录**：设置转换结果保存位置

4. **参数设置**（仅文档模式有效）：
   - **DPI**：图片清晰度，72（屏幕）~ 300（打印），默认 150
   - **文件名前缀**：输出图片的命名前缀，留空则用原文件名

5. 点击 **「开始转换」**

### 输出目录结构

文档转图片模式下，每个文件的输出会放在以原文件命名的子目录中：

```
输出目录/
├── 演示文稿A/
│   ├── 演示文稿A_01.png
│   ├── 演示文稿A_02.png
│   └── ...
├── 表格数据/
│   ├── 表格数据_01.jpg
│   └── ...
└── 报告文档.pdf          ← 文档→PDF 模式直接放在输出根目录
```

图片合并PDF模式下，所有图片合并为一个PDF文件：

```
输出目录/
└── 图片文件夹名.pdf       ← 合并后的PDF
```

---

## 支持的格式

### 输入（文档类）
- `.docx` — Word 文档
- `.xlsx` — Excel 表格
- `.pptx` — PowerPoint 演示文稿

### 输入（图片类）
- `.png` — 透明背景支持
- `.jpg` / `.jpeg`
- `.bmp`
- `.tiff` / `.tif`

### 输出
- `.pdf` — PDF 文档
- `.png` — 无损图片
- `.jpg` — 有损压缩（质量90%）

---

## 工作原理

```
┌─────────────────────────────────────────────────┐
│              文档 → 图片 流程                     │
│                                                  │
│  docx/xlsx/pptx ──→ LibreOffice ──→ PDF          │
│                                        │         │
│                                        ▼         │
│                                 PyMuPDF 渲染     │
│                                        │         │
│                                        ▼         │
│                                   PNG / JPG 序列 │
└─────────────────────────────────────────────────┘

┌─────────────────────────────────────────────────┐
│              图片 → PDF 流程                      │
│                                                  │
│  png/jpg/bmp ──→ PyMuPDF ──→ 合并为单个 PDF      │
└─────────────────────────────────────────────────┘
```

---

## 打包为 EXE

如果你想自己打包：

```bash
# 安装依赖
pip install pymupdf pyinstaller

# 一键打包（精简版，约30MB）
pyinstaller --onefile --windowed ^
  --name "批量转换工具" ^
  --exclude-module onnxruntime ^
  --exclude-module pymupdf_layout ^
  --exclude-module pymupdf4llm ^
  --exclude-module numpy ^
  --collect-submodules fitz ^
  batch_converter_gui.py
```

打包完成后，exe 在 `dist/` 目录下。

### 打包参数说明

| 参数 | 作用 |
|------|------|
| `--onefile` | 打包为单个 exe 文件 |
| `--windowed` | 不显示控制台黑窗口（GUI程序） |
| `--exclude-module xxx` | 排除无用模块，减小体积 |
| `--collect-submodules fitz` | 确保 PyMuPDF 完整打包 |

---

## 常见问题

### Q: 提示"未找到 LibreOffice"怎么办？
A: 请安装 LibreOffice，或在「LibreOffice」输入框中手动指定 `soffice.exe` 的路径。

### Q: 转换后的图片不清晰？
A: 提高 DPI 值（比如从 150 调到 200 或 300）。

### Q: 图片合并PDF时顺序不对？
A: 工具按**文件名排序**合并。请确保图片文件名是连续编号的（如 `page_01.png`, `page_02.png` ...）。

### Q: exe 文件为什么这么大（30MB）？
A: 因为打包了完整的 Python 运行时 + PyMuPDF PDF 引擎。已排除了 onnxruntime、numpy 等无用模块，体积已经是精简后的结果。

---

## 项目结构

```
doc-batch-converter/
├── batch_converter_gui.py   # 主程序（GUI + 全部转换逻辑）
├── README.md                # 本文件
└── .gitignore               # Git 忽略规则
```

---

## 技术栈

- **GUI**: tkinter（Python 内置）
- **文档转换**: LibreOffice headless 模式
- **PDF 处理**: PyMuPDF (fitz)
- **打包**: PyInstaller

---

## License

MIT

---

> 南无阿弥陀佛 🙏
