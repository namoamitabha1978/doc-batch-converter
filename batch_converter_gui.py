# -*- coding: utf-8 -*-
"""
批量转换工具
功能：
  1. 文档 → PDF / PNG序列 / JPG序列（docx / xlsx / pptx）
  2. 图片 → 合并为 PDF（png / jpg / jpeg / bmp / tiff）
依赖：LibreOffice (soffice.exe)、PyMuPDF (fitz)
"""

import os
import sys
import threading
import subprocess
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

try:
    import fitz  # PyMuPDF
except ImportError:
    fitz = None


# ──────────────────────────── 配置 ────────────────────────────
DEFAULT_SOFFICE = r"C:\Program Files\LibreOffice\program\soffice.exe"
DEFAULT_DPI = 150
DOC_EXTS = {".docx", ".xlsx", ".pptx"}
IMG_EXTS = {".png", ".jpg", ".jpeg", ".bmp", ".tiff", ".tif"}

# 输出格式选项
FMT_DOC_TO_PDF = "文档 → PDF"
FMT_DOC_TO_PNG = "文档 → PNG 序列"
FMT_DOC_TO_JPG = "文档 → JPG 序列"
FMT_IMG_TO_PDF = "图片 → 合并为 PDF"
OUTPUT_FORMATS = [FMT_DOC_TO_PDF, FMT_DOC_TO_PNG, FMT_DOC_TO_JPG, FMT_IMG_TO_PDF]


# ──────────────────────────── 核心转换逻辑 ────────────────────────────
def find_soffice(custom_path=""):
    """查找 soffice.exe"""
    if custom_path and os.path.isfile(custom_path):
        return custom_path
    if os.path.isfile(DEFAULT_SOFFICE):
        return DEFAULT_SOFFICE
    x86 = r"C:\Program Files (x86)\LibreOffice\program\soffice.exe"
    if os.path.isfile(x86):
        return x86
    return None


def collect_files(input_path, exts):
    """收集输入路径下指定扩展名的文件"""
    files = []
    if os.path.isfile(input_path):
        ext = os.path.splitext(input_path)[1].lower()
        if ext in exts:
            files.append(input_path)
    elif os.path.isdir(input_path):
        for root, _, names in os.walk(input_path):
            for n in names:
                ext = os.path.splitext(n)[1].lower()
                if ext in exts:
                    files.append(os.path.join(root, n))
    return sorted(files)


def doc_to_pdf(soffice_path, doc_path, output_dir, log_func):
    """用 LibreOffice 将文档转为 PDF"""
    os.makedirs(output_dir, exist_ok=True)
    profile = f"file:///C:/temp/lo_profile_{os.getpid()}_{abs(hash(doc_path)) % 100000}"
    cmd = [
        soffice_path,
        "--headless",
        f"-env:UserInstallation={profile}",
        "--convert-to", "pdf",
        "--outdir", output_dir,
        doc_path,
    ]
    log_func(f"  LibreOffice: {os.path.basename(doc_path)} → PDF")
    result = subprocess.run(cmd, capture_output=True, text=True, timeout=300)
    if result.returncode != 0:
        raise RuntimeError(f"LibreOffice 失败 (码 {result.returncode})")

    pdf_name = os.path.splitext(os.path.basename(doc_path))[0] + ".pdf"
    pdf_path = os.path.join(output_dir, pdf_name)
    if not os.path.isfile(pdf_path):
        raise FileNotFoundError(f"未生成 PDF: {pdf_name}")
    return pdf_path


def pdf_to_images(pdf_path, output_dir, dpi, img_format, prefix, log_func, cancel_check=None):
    """用 PyMuPDF 将 PDF 逐页转为图片"""
    doc = fitz.open(pdf_path)
    total = len(doc)
    zoom = dpi / 72
    mat = fitz.Matrix(zoom, zoom)
    base = prefix if prefix else os.path.splitext(os.path.basename(pdf_path))[0]
    ext = ".png" if img_format.upper() == "PNG" else ".jpg"

    success = 0
    for i in range(total):
        if cancel_check and cancel_check():
            break
        page = doc[i]
        pix = page.get_pixmap(matrix=mat)
        out_name = f"{base}_{i+1:02d}{ext}"
        out_path = os.path.join(output_dir, out_name)
        if img_format.upper() == "JPG":
            if pix.alpha:
                pix = fitz.Pixmap(pix, 0)
            pix.save(out_path, jpg_quality=90)
        else:
            pix.save(out_path)
        success += 1
    doc.close()
    return success, total


def images_to_pdf(image_paths, output_pdf_path, log_func, cancel_check=None):
    """将多张图片按顺序合并为一个 PDF"""
    doc = fitz.open()
    total = len(image_paths)
    success = 0

    for i, img_path in enumerate(image_paths):
        if cancel_check and cancel_check():
            break
        try:
            # 用图片创建新页面
            img = fitz.open(img_path)
            rect = img[0].rect
            page = doc.new_page(width=rect.width, height=rect.height)
            page.insert_image(rect, filename=img_path)
            img.close()
            success += 1
            if (i + 1) % 10 == 0 or i == total - 1:
                log_func(f"  已合并 {i+1}/{total} 张")
        except Exception as e:
            log_func(f"  跳过 {os.path.basename(img_path)}: {e}")

    doc.save(output_pdf_path)
    doc.close()
    return success, total


# ──────────────────────────── GUI 应用 ────────────────────────────
class BatchConverterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("批量转换工具 (文档↔图片) — 南无阿弥陀佛")
        self.root.geometry("740x640")
        self.root.resizable(True, True)

        self.input_path = tk.StringVar()
        self.out_dir = tk.StringVar()
        self.dpi_var = tk.IntVar(value=DEFAULT_DPI)
        self.prefix_var = tk.StringVar(value="")
        self.soffice_var = tk.StringVar(value=DEFAULT_SOFFICE)
        self.out_format_var = tk.StringVar(value=FMT_DOC_TO_PNG)
        self.is_converting = False
        self.cancelled = False

        self._build_ui()
        self._on_format_change()

    def _build_ui(self):
        pad = {"padx": 8, "pady": 4}

        # ── 输入 ──
        frm1 = ttk.LabelFrame(self.root, text=" 输入（文件或文件夹） ")
        frm1.pack(fill="x", **pad)
        row = ttk.Frame(frm1)
        row.pack(fill="x", padx=8, pady=6)
        ttk.Entry(row, textvariable=self.input_path).pack(side="left", fill="x", expand=True)
        ttk.Button(row, text="选择文件…", command=self._browse_file).pack(side="left", padx=(6, 2))
        ttk.Button(row, text="选择文件夹…", command=self._browse_folder).pack(side="left", padx=2)

        # ── 输出目录 ──
        frm2 = ttk.LabelFrame(self.root, text=" 输出目录 ")
        frm2.pack(fill="x", **pad)
        row = ttk.Frame(frm2)
        row.pack(fill="x", padx=8, pady=6)
        ttk.Entry(row, textvariable=self.out_dir).pack(side="left", fill="x", expand=True)
        ttk.Button(row, text="浏览…", command=self._browse_outdir).pack(side="left", padx=(6, 0))

        # ── 参数 ──
        frm3 = ttk.LabelFrame(self.root, text=" 参数设置 ")
        frm3.pack(fill="x", **pad)

        row1 = ttk.Frame(frm3)
        row1.pack(fill="x", padx=8, pady=6)
        ttk.Label(row1, text="转换模式:").pack(side="left")
        self.fmt_combo = ttk.Combobox(row1, textvariable=self.out_format_var, values=OUTPUT_FORMATS, state="readonly", width=18)
        self.fmt_combo.pack(side="left", padx=(4, 16))
        self.fmt_combo.bind("<<ComboboxSelected>>", lambda e: self._on_format_change())

        self.dpi_label = ttk.Label(row1, text="DPI:")
        self.dpi_label.pack(side="left")
        self.dpi_spin = ttk.Spinbox(row1, from_=72, to=300, increment=10, textvariable=self.dpi_var, width=8)
        self.dpi_spin.pack(side="left", padx=(4, 16))

        self.prefix_label = ttk.Label(row1, text="文件名前缀:")
        self.prefix_label.pack(side="left")
        self.prefix_entry = ttk.Entry(row1, textvariable=self.prefix_var, width=16)
        self.prefix_entry.pack(side="left", padx=(4, 0))

        row2 = ttk.Frame(frm3)
        row2.pack(fill="x", padx=8, pady=(0, 6))
        ttk.Label(row2, text="LibreOffice:").pack(side="left")
        ttk.Entry(row2, textvariable=self.soffice_var).pack(side="left", fill="x", expand=True, padx=(4, 0))
        ttk.Button(row2, text="…", width=3, command=self._browse_soffice).pack(side="left", padx=(4, 0))

        # ── 进度 ──
        frm4 = ttk.Frame(self.root)
        frm4.pack(fill="x", **pad)
        self.progress = ttk.Progressbar(frm4, mode="determinate")
        self.progress.pack(fill="x", padx=8, pady=(4, 0))
        self.progress_label = ttk.Label(frm4, text="就绪", anchor="center")
        self.progress_label.pack(fill="x")

        # ── 按钮 ──
        frm5 = ttk.Frame(self.root)
        frm5.pack(fill="x", **pad)
        self.btn_start = ttk.Button(frm5, text="开始转换", command=self._start_convert)
        self.btn_start.pack(side="left", padx=8)
        self.btn_cancel = ttk.Button(frm5, text="取消", command=self._cancel_convert, state="disabled")
        self.btn_cancel.pack(side="left", padx=4)
        self.lbl_count = ttk.Label(frm5, text="", foreground="#666")
        self.lbl_count.pack(side="right", padx=8)

        # ── 日志 ──
        frm6 = ttk.LabelFrame(self.root, text=" 日志 ")
        frm6.pack(fill="both", expand=True, **pad)
        self.log_text = tk.Text(frm6, height=10, wrap="word", state="disabled", bg="#f8f8f8")
        scroll = ttk.Scrollbar(frm6, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=scroll.set)
        scroll.pack(side="right", fill="y")
        self.log_text.pack(fill="both", expand=True, padx=8, pady=6)

    def _on_format_change(self):
        """根据转换模式显示/隐藏DPI和前缀设置"""
        fmt = self.out_format_var.get()
        is_img_mode = (fmt == FMT_IMG_TO_PDF)
        # 图片转PDF模式下，DPI和前缀无意义
        state = "disabled" if is_img_mode else "normal"
        self.dpi_label.configure(state=state)
        self.dpi_spin.configure(state=state)
        self.prefix_label.configure(state=state)
        self.prefix_entry.configure(state=state)

    # ──────────────────────── 事件 ────────────────────────
    def _browse_file(self):
        fmt = self.out_format_var.get()
        if fmt == FMT_IMG_TO_PDF:
            ftypes = [("图片文件", "*.png *.jpg *.jpeg *.bmp *.tiff"), ("所有文件", "*.*")]
            title = "选择图片文件"
        else:
            ftypes = [("文档文件", "*.docx *.xlsx *.pptx"), ("所有文件", "*.*")]
            title = "选择文档"
        path = filedialog.askopenfilename(title=title, filetypes=ftypes)
        if path:
            self.input_path.set(path)
            if not self.out_dir.get():
                d = os.path.join(os.path.dirname(path), "输出")
                self.out_dir.set(d)

    def _browse_folder(self):
        d = filedialog.askdirectory(title="选择包含文件的文件夹")
        if d:
            self.input_path.set(d)
            if not self.out_dir.get():
                self.out_dir.set(os.path.join(d, "输出"))

    def _browse_outdir(self):
        d = filedialog.askdirectory(title="选择输出目录")
        if d:
            self.out_dir.set(d)

    def _browse_soffice(self):
        path = filedialog.askopenfilename(
            title="选择 soffice.exe",
            filetypes=[("soffice.exe", "soffice.exe"), ("所有文件", "*.*")],
        )
        if path:
            self.soffice_var.set(path)

    def _log(self, msg):
        self.log_text.configure(state="normal")
        self.log_text.insert("end", msg + "\n")
        self.log_text.see("end")
        self.log_text.configure(state="disabled")
        self.root.update_idletasks()

    def _set_progress(self, cur, total, label=""):
        self.progress["maximum"] = total
        self.progress["value"] = cur
        self.progress_label.configure(text=f"{cur} / {total}  {label}")
        self.root.update_idletasks()

    def _cancel_convert(self):
        self.cancelled = True
        self._log("正在取消…")

    def _start_convert(self):
        if self.is_converting:
            return

        inp = self.input_path.get().strip()
        outdir = self.out_dir.get().strip()
        dpi = self.dpi_var.get()
        prefix = self.prefix_var.get().strip()
        soffice = self.soffice_var.get().strip()
        out_fmt = self.out_format_var.get()

        if not inp or not os.path.exists(inp):
            messagebox.showerror("错误", "请选择有效的输入文件或文件夹")
            return
        if not outdir:
            messagebox.showerror("错误", "请选择输出目录")
            return
        if fitz is None:
            messagebox.showerror("错误", "PyMuPDF 未安装")
            return

        is_img_mode = (out_fmt == FMT_IMG_TO_PDF)
        target_exts = IMG_EXTS if is_img_mode else DOC_EXTS

        files = collect_files(inp, target_exts)
        if not files:
            ext_name = "图片" if is_img_mode else "docx / xlsx / pptx"
            messagebox.showinfo("提示", f"未找到任何 {ext_name} 文件")
            return

        # 文档模式需要LibreOffice
        soffice_found = None
        if not is_img_mode:
            soffice_found = find_soffice(soffice)
            if not soffice_found:
                messagebox.showerror("错误", "未找到 LibreOffice (soffice.exe)")
                return

        self.lbl_count.configure(text=f"共 {len(files)} 个文件")
        self.is_converting = True
        self.cancelled = False
        self.btn_start.configure(state="disabled")
        self.btn_cancel.configure(state="normal")
        self.progress["value"] = 0
        self._log("=" * 55)
        self._log(f"输入: {inp}")
        self._log(f"输出: {outdir}")
        self._log(f"模式: {out_fmt}  文件数: {len(files)}")

        t = threading.Thread(
            target=self._convert_worker,
            args=(files, outdir, dpi, prefix, soffice_found, out_fmt, is_img_mode),
            daemon=True,
        )
        t.start()

    def _convert_worker(self, files, outdir, dpi, prefix, soffice, out_fmt, is_img_mode):
        try:
            os.makedirs(outdir, exist_ok=True)
            total = len(files)
            done = 0

            if is_img_mode:
                # ── 图片合并为PDF ──
                self._log(f"\n开始合并 {total} 张图片为一个PDF...")
                # 输出PDF名：用输入文件夹名或第一个文件名
                if os.path.isdir(self.input_path.get().strip()):
                    pdf_name = os.path.basename(self.input_path.get().strip()) + ".pdf"
                else:
                    pdf_name = os.path.splitext(os.path.basename(files[0]))[0] + "_合并.pdf"
                if prefix:
                    pdf_name = prefix + "_" + pdf_name
                output_pdf = os.path.join(outdir, pdf_name)

                count, _ = images_to_pdf(
                    files, output_pdf, self._log,
                    cancel_check=lambda: self.cancelled,
                )
                size_kb = os.path.getsize(output_pdf) // 1024
                self._set_progress(total, total, "完成")
                self._log(f"\n✓ 合并完成！{count} 张图片 → {pdf_name} ({size_kb} KB)")
                messagebox.showinfo("完成", f"图片合并完成！\n{count} 张图片\n输出:\n{output_pdf}")
                return

            # ── 文档转换模式 ──
            total_pages = 0
            for idx, fpath in enumerate(files):
                if self.cancelled:
                    break

                fname = os.path.basename(fpath)
                self._log(f"\n[{idx+1}/{total}] {fname}")
                self._set_progress(idx, total, fname)

                sub_name = os.path.splitext(fname)[0]
                file_out_dir = os.path.join(outdir, sub_name)

                if out_fmt == FMT_DOC_TO_PDF:
                    pdf_path = doc_to_pdf(soffice, fpath, outdir, self._log)
                    size_kb = os.path.getsize(pdf_path) // 1024
                    self._log(f"  ✓ PDF 完成 ({size_kb} KB)")
                else:
                    tmp_dir = os.path.join(file_out_dir, "_tmp")
                    os.makedirs(tmp_dir, exist_ok=True)
                    pdf_path = doc_to_pdf(soffice, fpath, tmp_dir, self._log)

                    img_fmt = "PNG" if "PNG" in out_fmt else "JPG"
                    img_prefix = prefix if prefix else sub_name
                    count, pages = pdf_to_images(
                        pdf_path, file_out_dir, dpi, img_fmt, img_prefix,
                        log_func=self._log, cancel_check=lambda: self.cancelled,
                    )
                    total_pages += count
                    self._log(f"  ✓ {count} 页 {img_fmt} 已导出")

                    try:
                        os.remove(pdf_path)
                        os.rmdir(tmp_dir)
                    except OSError:
                        pass

                done += 1

            self._set_progress(total, total, "完成")
            if not self.cancelled:
                extra = f"，共 {total_pages} 张图片" if out_fmt != FMT_DOC_TO_PDF else ""
                self._log(f"\n{'='*55}")
                self._log(f"✓ 全部完成！成功转换 {done}/{total} 个文件{extra}")
                messagebox.showinfo("完成", f"转换完成！\n成功 {done}/{total} 个文件{extra}\n输出目录:\n{outdir}")
            else:
                self._log("\n已取消")

        except Exception as e:
            self._log(f"✗ 错误: {e}")
            messagebox.showerror("转换失败", str(e))
        finally:
            self.is_converting = False
            self.btn_start.configure(state="normal")
            self.btn_cancel.configure(state="disabled")


# ──────────────────────────── 启动 ────────────────────────────
if __name__ == "__main__":
    root = tk.Tk()
    app = BatchConverterApp(root)
    root.mainloop()
