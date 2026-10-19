import subprocess
import shutil
import platform
import os
from pathlib import Path
from typing import Optional

class PdfRouter:
    @classmethod
    def detect_engine(cls) -> Optional[str]:
        """Detect the available PDF conversion engine following priority: office > libreoffice > chrome."""
        if cls.is_office_available():
            return "office"
        if cls.is_libreoffice_available():
            return "libreoffice"
        if cls.is_chrome_available():
            return "chrome"
        return None

    @classmethod
    def is_office_available(cls) -> bool:
        """Check if Microsoft Word or WPS Office is installed and available."""
        if platform.system().lower() != "windows":
            return False
        try:
            import winreg
            for progid in ["Word.Application", "KWPS.Application", "WPS.Application"]:
                try:
                    with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, progid):
                        return True
                except Exception:
                    pass
            for app in ["winword.exe", "wps.exe"]:
                try:
                    with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, rf"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\{app}"):
                        return True
                except Exception:
                    pass
        except Exception:
            pass
        return bool(shutil.which("winword") or shutil.which("wps"))

    @classmethod
    def is_libreoffice_available(cls) -> bool:
        """Check if LibreOffice soffice CLI is installed."""
        return cls.find_libreoffice_bin() is not None

    @classmethod
    def is_chrome_available(cls) -> bool:
        """Check if Google Chrome or Microsoft Edge headless is installed."""
        return cls.find_chrome_or_edge_bin() is not None

    @classmethod
    def find_libreoffice_bin(cls) -> Optional[str]:
        soffice = shutil.which("soffice") or shutil.which("libreoffice")
        if soffice:
            return soffice
        
        candidates = [
            r"C:\Program Files\LibreOffice\program\soffice.exe",
            r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
            "/usr/bin/soffice",
            "/usr/bin/libreoffice",
            "/Applications/LibreOffice.app/Contents/MacOS/soffice"
        ]
        for c in candidates:
            if os.path.exists(c):
                return c
        return None

    @classmethod
    def find_chrome_or_edge_bin(cls) -> Optional[str]:
        # Check Windows registry first for installed Chrome/Edge
        if platform.system().lower() == "windows":
            try:
                import winreg
                for app in ["chrome.exe", "msedge.exe"]:
                    try:
                        key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, rf"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\{app}")
                        path, _ = winreg.QueryValueEx(key, "")
                        if path and os.path.exists(path):
                            return path
                    except Exception:
                        pass
            except Exception:
                pass

        # Check PATH
        for bin_name in ["google-chrome", "chrome", "chromium", "chromium-browser", "msedge"]:
            found = shutil.which(bin_name)
            if found:
                return found

        # Check standard file system paths
        candidates = [
            r"C:\Program Files\Google\Chrome\Application\chrome.exe",
            r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
            os.path.expandvars(r"%LOCALAPPDATA%\Google\Chrome\Application\chrome.exe"),
            r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
            r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
            "/usr/bin/google-chrome",
            "/usr/bin/chromium",
            "/usr/bin/chromium-browser",
            "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
            "/Applications/Microsoft Edge.app/Contents/MacOS/Microsoft Edge"
        ]
        for c in candidates:
            if c and os.path.exists(c):
                return c
        return None

    @classmethod
    def convert_docx_to_pdf(cls, docx_path: Path, pdf_path: Path, force_engine: str = "auto") -> bool:
        """
        Convert DOCX to PDF using the designated or auto-detected engine.
        Priority in 'auto' mode: office > libreoffice > chrome.
        """
        docx_path = Path(docx_path).resolve()
        pdf_path = Path(pdf_path).resolve()
        pdf_path.parent.mkdir(parents=True, exist_ok=True)

        if force_engine == "office":
            return cls._try_office_convert(docx_path, pdf_path)
        if force_engine == "libreoffice":
            return cls._try_libreoffice_convert(docx_path, pdf_path)
        if force_engine == "chrome":
            return cls._try_chrome_convert(docx_path, pdf_path)
        if force_engine == "univer":
            return cls._try_univer_convert(docx_path, pdf_path)

        # Priority: office > libreoffice > chrome
        if cls._try_office_convert(docx_path, pdf_path):
            return True

        if cls._try_libreoffice_convert(docx_path, pdf_path):
            return True

        if cls._try_chrome_convert(docx_path, pdf_path):
            return True

        # Last resort Univer adapter
        if cls._try_univer_convert(docx_path, pdf_path):
            return True

        return False

    @classmethod
    def _try_office_convert(cls, docx_path: Path, pdf_path: Path) -> bool:
        if platform.system().lower() != "windows":
            return False
        
        # Try MS Word via COM
        try:
            import win32com.client
            word = win32com.client.Dispatch("Word.Application")
            try:
                word.Visible = 0
                doc = word.Documents.Open(str(docx_path), ReadOnly=1)
                doc.ExportAsFixedFormat(str(pdf_path), 17) # 17 = wdExportFormatPDF
                doc.Close(SaveChanges=0)
                if pdf_path.exists() and pdf_path.stat().st_size > 0:
                    return True
            finally:
                word.Quit()
        except Exception:
            pass

        # Try WPS Office via COM
        for wps_progid in ["KWPS.Application", "WPS.Application"]:
            try:
                import win32com.client
                wps = win32com.client.Dispatch(wps_progid)
                try:
                    wps.Visible = 0
                    doc = wps.Documents.Open(str(docx_path), ReadOnly=1)
                    doc.ExportAsFixedFormat(str(pdf_path), 17)
                    doc.Close(SaveChanges=0)
                    if pdf_path.exists() and pdf_path.stat().st_size > 0:
                        return True
                finally:
                    wps.Quit()
            except Exception:
                pass

        return False

    @classmethod
    def _try_libreoffice_convert(cls, docx_path: Path, pdf_path: Path) -> bool:
        soffice = cls.find_libreoffice_bin()
        if not soffice:
            return False

        try:
            outdir = str(pdf_path.parent)
            cmd = [soffice, "--headless", "--convert-to", "pdf", "--outdir", outdir, str(docx_path)]
            proc = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=60)
            
            generated_pdf = pdf_path.parent / (docx_path.stem + ".pdf")
            if generated_pdf.exists() and generated_pdf != pdf_path:
                shutil.move(str(generated_pdf), str(pdf_path))

            return pdf_path.exists() and pdf_path.stat().st_size > 0
        except Exception:
            return False

    @classmethod
    def _try_chrome_convert(cls, docx_path: Path, pdf_path: Path) -> bool:
        chrome_bin = cls.find_chrome_or_edge_bin()
        if not chrome_bin:
            return False

        temp_html = pdf_path.with_suffix(".tmp.html")
        try:
            import mammoth
            with open(docx_path, "rb") as docx_file:
                result = mammoth.convert_to_html(docx_file)
                html_body = result.value

            full_html = (
                "<!DOCTYPE html>\n<html>\n<head>\n<meta charset=\"utf-8\">\n"
                "<style>\n"
                "  @page { size: A4; margin: 15mm 15mm 15mm 15mm; }\n"
                "  body { font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, 'SimSun', 'Microsoft YaHei', sans-serif; font-size: 10.5pt; line-height: 1.5; color: #111827; margin: 0; padding: 0; }\n"
                "  table { border-collapse: collapse; width: 100%; margin-bottom: 12px; page-break-inside: auto; }\n"
                "  tr { page-break-inside: avoid; page-break-after: auto; }\n"
                "  th, td { border: 1px solid #d1d5db; padding: 6px 10px; text-align: left; font-size: 9.5pt; }\n"
                "  th { background-color: #f3f4f6; font-weight: 600; }\n"
                "  img { max-width: 100%; height: auto; }\n"
                "  h1, h2, h3, h4, h5 { page-break-after: avoid; color: #1f2937; }\n"
                "  h1 { font-size: 16pt; margin: 12px 0 6px 0; }\n"
                "  h2 { font-size: 13pt; margin: 10px 0 5px 0; }\n"
                "  p { margin: 4px 0; }\n"
                "</style>\n</head>\n<body>\n"
                + html_body +
                "\n</body>\n</html>"
            )
            temp_html.write_text(full_html, encoding="utf-8")

            cmd = [
                str(chrome_bin),
                "--headless=new",
                "--disable-gpu",
                "--no-sandbox",
                "--no-pdf-header-footer",
                f"--print-to-pdf={pdf_path}",
                f"file:///{temp_html.resolve().as_posix()}"
            ]
            proc = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=60)
            return pdf_path.exists() and pdf_path.stat().st_size > 0
        except Exception:
            return False
        finally:
            if temp_html.exists():
                try:
                    temp_html.unlink()
                except Exception:
                    pass

    @staticmethod
    def _try_univer_convert(docx_path: Path, pdf_path: Path) -> bool:
        node_bin = shutil.which("node")
        if not node_bin:
            return False

        adapter_script = Path(__file__).resolve().parent.parent / "univer_adapter" / "univer-convert.js"
        if not adapter_script.exists():
            return False

        try:
            cmd = [node_bin, str(adapter_script), str(docx_path), str(pdf_path)]
            proc = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=30)
            return proc.returncode == 0 and pdf_path.exists() and pdf_path.stat().st_size > 0
        except Exception:
            return False
