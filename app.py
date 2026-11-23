import os
import shutil
import subprocess
import sys
import threading
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from PIL import Image

try:
    import fitz
except ImportError:
    fitz = None

SUPPORTED_INPUTS = {
    ".pdf", ".png", ".jpg", ".jpeg", ".bmp", ".tif", ".tiff", ".webp",
    ".doc", ".docx", ".odt", ".rtf", ".txt", ".html", ".htm",
    ".xls", ".xlsx", ".ods", ".csv",
    ".ppt", ".pptx", ".odp",
}
OUTPUT_FORMATS = ["PDF", "DOCX", "DOC", "ODT", "TXT", "HTML", "RTF", "XLSX", "ODS", "CSV", "PPTX", "ODP", "PNG", "JPG"]


class ConverterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("文档转换器")
        self.root.geometry("920x620")
        self.root.minsize(780, 520)

        self.files = []
        self.output_dir = tk.StringVar()

        self._build_ui()

    def _build_ui(self):
        top = ttk.Frame(self.root, padding=12)
        top.pack(fill="x")
        ttk.Button(top, text="添加文件", command=self.add_files).pack(side="left")
        ttk.Button(top, text="清空列表", command=self.clear_files).pack(side="left", padx=(8, 0))

        self.file_list = tk.Listbox(self.root, height=12)
        self.file_list.pack(fill="both", expand=True, padx=12)

        controls = ttk.LabelFrame(self.root, text="转换设置", padding=12)
        controls.pack(fill="x", padx=12, pady=(10, 0))

        ttk.Label(controls, text="目标格式").grid(row=0, column=0, sticky="w")
        self.output_format = tk.StringVar(value="PDF")
        self.format_combo = ttk.Combobox(controls, textvariable=self.output_format,
                                         values=OUTPUT_FORMATS, state="readonly", width=12)
        self.format_combo.grid(row=0, column=1, padx=(8, 20))

        ttk.Label(controls, text="输出目录").grid(row=0, column=2, sticky="w")
        ttk.Entry(controls, textvariable=self.output_dir).grid(row=0, column=3, sticky="ew", padx=(8, 8))
        ttk.Button(controls, text="选择", command=self.choose_output).grid(row=0, column=4)
        controls.columnconfigure(3, weight=1)

        bottom = ttk.Frame(self.root, padding=12)
        bottom.pack(fill="x")
        self.progress = ttk.Progressbar(bottom, mode="determeterminate")
        self.progress.pack(side="left", fill="x", expand=True)
        ttk.Button(bottom, text="开始转换", command=self.start_conversion).pack(side="right", padx=(12, 0))

        self.status = tk.StringVar(value="就绪")
        ttk.Label(self.root, textvariable=self.status, anchor="w").pack(fill="x", padx=12, pady=(0, 8))

    def add_files(self):
        selected = filedialog.askopenfilenames(
            title="选择要转换的文件",
            filetypes=[("支持的文档", " ".join("*" + ext for ext in sorted(SUPPORTED_INPUTS))), ("所有文件", "*.*")],
        )
        for name in selected:
            path = Path(name)
            if path.suffix.lower() in SUPPORTED_INPUTS and str(path) not in self.files:
                self.files.append(str(path))
        self.refresh_list()

    def clear_files(self):
        self.files.clear()
        self.refresh_list()

    def refresh_list(self):
        self.file_list.delete(0, tk.END)
        for path in self.files:
            self.file_list.insert(tk.END, path)
        self.status.set(f"已选择 {len(self.files)} 个文件")

    def choose_output(self):
        directory = filedialog.askdirectory(title="选择输出目录")
        if directory:
            self.output_dir.set(directory)

    def find_libreoffice(self):
        candidates = []
        if getattr(sys, "frozen", False):
            bundle = Path(sys._MEIPASS)
            candidates += [
                bundle / "LibreOffice" / "program" / "soffice.exe",
                bundle.parent / "LibreOffice" / "program" / "soffice.exe",
                Path.cwd() / "LibreOffice" / "program" / "soffice.exe",
            ]
        candidates += [
            Path(os.environ.get("PROGRAMFILES", "C:/Program Files")) / "LibreOffice" / "program" / "soffice.exe",
            Path(os.environ.get("PROGRAMFILES(X86)", "C:/Program Files (x86)")) / "LibreOffice" / "program" / "soffice.exe",
        ]
        for path in candidates:
            if path.is_file():
                return str(path)
        return shutil.which("soffice") or shutil.which("libreoffice")

    def convert_office(self, source: Path, target_format: str, output_dir: Path, soffice: str):
        profile = output_dir / ".lo-profile"
        profile.mkdir(parents=True, exist_ok=True)
        command = [
            soffice,
            f"-env:UserInstallation=file:///{profile.as_posix().lstrip('/')}",
            "--headless", "--convert-to", target_format.lower(), "--outdir", str(output_dir), str(source),
        ]
        result = subprocess.run(command, capture_output=True, text=True, timeout=180)
        if result.returncode != 0:
            raise RuntimeError(result.stderr.strip() or "LibreOffice conversion failed")

    def pdf_to_images(self, source: Path, target_format: str, output_dir: Path):
        if fitz is None:
            raise RuntimeError("PDF conversion library is missing")
        document = fitz.open(source)
        prefix = source.stem
        count = 0
        extension = "png" if target_format == "PNG" else "jpg"
        for page_number, page in enumerate(document):
            pixmap = page.get_pixmap(matrix=fitz.Matrix(2, 2), alpha=False)
            output = output_dir / f"{prefix}-page-{page_number + 1}.{extension}"
            pixmap.save(output)
            count += 1
        document.close()
        return count

    def image_to_pdf(self, source: Path, output_dir: Path):
        output = output_dir / f"{source.stem}.pdf"
        image = Image.open(source)
        if image.mode not in ("RGB", "L"):
            image = image.convert("RGB")
        image.save(output, "PDF", resolution=150.0)
        return 1

    def pdf_to_text(self, source: Path, output_dir: Path):
        if fitz is None:
            raise RuntimeError("PDF conversion library is missing")
        document = fitz.open(source)
        text = "\n\n".join(page.get_text("text") for page in document)
        document.close()
        output = output_dir / f"{source.stem}.txt"
        output.write_text(text, encoding="utf-8")
        return 1

    def convert_one(self, source: Path, target_format: str, output_dir: Path):
        suffix = source.suffix.lower()
        if suffix in {".png", ".jpg", ".jpeg", ".bmp", ".tif", ".tiff", ".webp"}:
            if target_format == "PDF":
                return self.image_to_pdf(source, output_dir)
        if suffix == ".pdf":
            if target_format in {"PNG", "JPG"}:
                return self.pdf_to_images(source, target_format, output_dir)
            if target_format == "TXT":
                return self.pdf_to_text(source, output_dir)

        soffice = self.find_libreoffice()
        if not soffice:
            raise RuntimeError("未找到 LibreOffice。请安装 LibreOffice 后重试。")
        self.convert_office(source, target_format, output_dir, soffice)
        return 1

    def start_conversion(self):
        if not self.files:
            messagebox.showwarning("文档转换器", "请先添加要转换的文件")
            return
        output_dir = Path(self.output_dir.get().strip()) if self.output_dir.get().strip() else None
        if output_dir is None:
            output_dir = Path(filedialog.askdirectory(title="选择输出目录"))
            if not str(output_dir):
                return
            self.output_dir.set(str(output_dir))
        output_dir.mkdir(parents=True, exist_ok=True)

        thread = threading.Thread(target=self.run_conversion, args=(output_dir,), daemon=True)
        thread.start()

    def run_conversion(self, output_dir: Path):
        target_format = self.output_format.get()
        total = len(self.files)
        failures = []
        successful = 0
        self.progress.configure(maximum=total)

        for index, name in enumerate(self.files, start=1):
            source = Path(name)
            self.status.set(f"正在转换：{source.name} ({index}/{total})")
            self.progress.configure(value=index - 1)
            try:
                successful += self.convert_one(source, target_format, output_dir)
            except Exception as error:
                failures.append(f"{source.name}: {error}")
            self.progress.configure(value=index)
            self.root.update_idletasks()

        self.status.set(f"转换完成：成功 {successful}，失败 {len(failures)}")
        if failures:
            messagebox.showerror("转换结果", "\n".join(failures[:20]))
        else:
            messagebox.showinfo("转换完成", f"已转换 {successful} 个文件\n输出目录：{output_dir}")


def main():
    root = tk.Tk()
    app = ConverterApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
