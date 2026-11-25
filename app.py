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
        self.root.title("DocConverter - 文档转换工作台")
        self.root.geometry("1080x720")
        self.root.minsize(900, 620)

        self.files = []
        self.output_dir = tk.StringVar()
        self.status = tk.StringVar(value="就绪")
        self.output_format = tk.StringVar(value="PDF")
        self.count_text = tk.StringVar(value="0 个文件")

        self.configure_theme()
        self._build_ui()

    def configure_theme(self):
        style = ttk.Style(self.root)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass

        style.configure(".", background="#F5F6F8", foreground="#202124",
                        font=("Microsoft YaHei UI", 10))
        style.configure("TFrame", background="#F5F6F8")
        style.configure("Canvas.TFrame", background="#FFFFFF")
        style.configure("Main.TFrame", background="#F5F6F8")
        style.configure("Header.TFrame", background="#FFFFFF")
        style.configure("TLabel", background="#FFFFFF", foreground="#5F6368")
        style.configure("Title.TLabel", background="#FFFFFF", foreground="#202124",
                        font=("Microsoft YaHei UI", 18, "bold"))
        style.configure("Muted.TLabel", background="#FFFFFF", foreground="#8A9099",
                        font=("Microsoft YaHei UI", 9))
        style.configure("Section.TLabel", background="#FFFFFF", foreground="#202124",
                        font=("Microsoft YaHei UI", 11, "bold"))
        style.configure("Status.TLabel", background="#FFFFFF", foreground="#8A9099",
                        font=("Microsoft YaHei UI", 9))
        style.configure("TButton", padding=(12, 8), relief="flat")
        style.map("TButton", background=[("active", "#EEF1F4")])
        style.configure("Primary.TButton", background="#202124", foreground="#FFFFFF",
                        font=("Microsoft YaHei UI", 10, "bold"), padding=(17, 10), relief="flat")
        style.map("Primary.TButton", background=[("active", "#34373A")])
        style.configure("Secondary.TButton", background="#FFFFFF", foreground="#202124",
                        borderwidth=1, relief="solid", padding=(12, 8))
        style.map("Secondary.TButton", background=[("active", "#F2F4F5")])
        style.configure("TCombobox", fieldbackground="#FFFFFF", background="#FFFFFF",
                        foreground="#202124", padding=8)
        style.configure("Horizontal.TProgressbar", background="#202124",
                        troughcolor="#ECEFF1", thickness=6)

    def _build_ui(self):
        self.root.configure(background="#F5F6F8")

        shell = tk.Frame(self.root, bg="#FFFFFF", highlightbackground="#DDE1E5",
                         highlightthickness=1)
        shell.pack(fill="both", expand=True, padx=18, pady=18)

        header = tk.Frame(shell, bg="#FFFFFF", height=78)
        header.pack(fill="x", padx=24)
        header.pack_propagate(False)

        title_block = tk.Frame(header, bg="#FFFFFF")
        title_block.pack(side="left", fill="y")
        tk.Label(title_block, text="DocConverter", bg="#FFFFFF", fg="#202124",
                 font=("Microsoft YaHei UI", 17, "bold")).pack(anchor="w", pady=(16, 0))
        tk.Label(title_block, text="文档转换工作台", bg="#FFFFFF", fg="#8A9099",
                 font=("Microsoft YaHei UI", 9)).pack(anchor="w", pady=(2, 0))

        header_actions = tk.Frame(header, bg="#FFFFFF")
        header_actions.pack(side="right", fill="y")
        ttk.Button(header_actions, text="清空", style="Secondary.TButton",
                   command=self.clear_files).pack(side="left", pady=22)
        ttk.Button(header_actions, text="移除选中", style="Secondary.TButton",
                   command=self.remove_selected).pack(side="left", padx=(8, 0), pady=22)

        divider = tk.Frame(shell, bg="#ECEFF1", height=1)
        divider.pack(fill="x")

        body = tk.Frame(shell, bg="#FFFFFF")
        body.pack(fill="both", expand=True, padx=24, pady=20)

        files_section = tk.Frame(body, bg="#FFFFFF")
        files_section.pack(side="left", fill="both", expand=True)

        files_head = tk.Frame(files_section, bg="#FFFFFF", height=44)
        files_head.pack(fill="x")
        files_head.pack_propagate(False)
        tk.Label(files_head, text="待转换文件", bg="#FFFFFF", fg="#202124",
                 font=("Microsoft YaHei UI", 12, "bold")).pack(side="left")
        tk.Label(files_head, textvariable=self.count_text, bg="#FFFFFF", fg="#8A9099",
                 font=("Microsoft YaHei UI", 9)).pack(side="right")

        list_frame = tk.Frame(files_section, bg="#F8F9FA", highlightbackground="#DDE1E5",
                              highlightthickness=1)
        list_frame.pack(fill="both", expand=True)
        scrollbar = tk.Scrollbar(list_frame)
        scrollbar.pack(side="right", fill="y")
        self.file_list = tk.Listbox(
            list_frame,
            height=12,
            borderwidth=0,
            highlightthickness=0,
            selectbackground="#E7EBEF",
            selectforeground="#202124",
            background="#F8F9FA",
            foreground="#3C4043",
            activestyle="none",
            font=("Microsoft YaHei UI", 10),
            yscrollcommand=scrollbar.set,
        )
        self.file_list.pack(side="left", fill="both", expand=True, padx=(8, 0), pady=8)
        scrollbar.config(command=self.file_list.yview)

        add_row = tk.Frame(files_section, bg="#FFFFFF", height=54)
        add_row.pack(fill="x")
        add_row.pack_propagate(False)
        ttk.Button(add_row, text="＋  添加文件", style="Secondary.TButton",
                   command=self.add_files).pack(side="left", pady=10)
        tk.Label(add_row, text="支持 PDF、图片及 Office 文档",
                 bg="#FFFFFF", fg="#A0A5AC",
                 font=("Microsoft YaHei UI", 9)).pack(side="right", pady=10)

        settings = tk.Frame(body, bg="#FFFFFF", width=268)
        settings.pack(side="right", fill="y", padx=(24, 0))
        settings.pack_propagate(False)

        tk.Label(settings, text="转换设置", bg="#FFFFFF", fg="#202124",
                 font=("Microsoft YaHei UI", 12, "bold")).pack(anchor="w", pady=(0, 18))

        tk.Label(settings, text="目标格式", bg="#FFFFFF", fg="#5F6368",
                 font=("Microsoft YaHei UI", 9)).pack(anchor="w")
        self.format_combo = ttk.Combobox(settings, textvariable=self.output_format,
                                         values=OUTPUT_FORMATS, state="readonly", width=24)
        self.format_combo.pack(anchor="w", fill="x", pady=(7, 20), ipady=5)

        tk.Label(settings, text="输出目录", bg="#FFFFFF", fg="#5F6368",
                 font=("Microsoft YaHei UI", 9)).pack(anchor="w")
        self.output_entry = tk.Entry(settings, textvariable=self.output_dir, relief="solid",
                                     borderwidth=1, bg="#FFFFFF", fg="#3C4043",
                                     font=("Microsoft YaHei UI", 9), highlightthickness=0)
        self.output_entry.pack(anchor="w", fill="x", pady=(7, 8), ipady=7)
        ttk.Button(settings, text="选择输出目录", style="Secondary.TButton",
                   command=self.choose_output).pack(anchor="w", fill="x")

        convert_button = tk.Frame(settings, bg="#FFFFFF")
        convert_button.pack(side="bottom", fill="x")
        ttk.Button(convert_button, text="开始转换", style="Primary.TButton",
                   command=self.start_conversion).pack(fill="x", ipady=4)

        footer = tk.Frame(shell, bg="#FFFFFF", height=52)
        footer.pack(fill="x", padx=24)
        footer.pack_propagate(False)
        tk.Frame(footer, bg="#ECEFF1", height=1).pack(fill="x", side="top")
        self.progress = ttk.Progressbar(footer, mode="determinate")
        self.progress.pack(side="left", fill="x", expand=True, pady=(16, 14))
        ttk.Label(footer, textvariable=self.status, style="Status.TLabel",
                  width=36, anchor="e").pack(side="right", padx=(16, 0), pady=(14, 12))

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
            name = Path(path).name
            suffix = Path(path).suffix.upper().lstrip(".") or "文件"
            size = 0
            try:
                size = Path(path).stat().st_size
            except OSError:
                pass
            if size < 1024:
                size_text = f"{size} B"
            elif size < 1024 * 1024:
                size_text = f"{size / 1024:.1f} KB"
            else:
                size_text = f"{size / 1024 / 1024:.1f} MB"
            self.file_list.insert(tk.END, f"  {suffix:<5}  {name:<55}  {size_text}")
        self.count_text.set(f"{len(self.files)} 个文件")
        self.status.set(f"已选择 {len(self.files)} 个文件")

    def remove_selected(self):
        selected = list(self.file_list.curselection())
        for index in reversed(selected):
            del self.files[index]
        self.refresh_list()

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
