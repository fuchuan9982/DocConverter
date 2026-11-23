# DocConverter

Windows 桌面文档转换工具，支持批量转换 PDF、图片和常见 Office 文档。

## 支持格式

- 图片：PNG、JPG/JPEG、BMP、TIFF、WebP
- PDF：转 PNG/JPG、TXT，以及通过 LibreOffice 转 DOCX/DOC/ODT/RTF 等
- Office：DOC/DOCX/ODT/RTF、XLS/XLSX/ODS/CSV、PPT/PPTX/ODP
- 文本：TXT、HTML/HTM

## Windows 枚举

1. 打开 GitHub Actions。
2. 选择 `Build DocConverter Windows EXE`。
3. 下载 `DocConverter-Windows-x64` artifact。
4. 解压并运行 `DocConverter.exe`。

GitHub Actions 会将 LibreOffice 随应用打包，因此目标 Windows 电脑无需额外安装 Office 或 LibreOffice。

## 本地运行

```powershell
python -m pip install -r requirements.txt
python app.py
```

PDF 与图片转换使用 Pillow 和 PyMuPDF；Office、文本、网页文档转换使用 LibreOffice。
