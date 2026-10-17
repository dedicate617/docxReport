import subprocess
import shutil
import platform
import os
from pathlib import Path

class PdfRouter:
    @classmethod
    def convert_docx_to_pdf(cls, docx_path: Path, pdf_path: Path, force_engine: str = "auto") -> bool:
        docx_path = Path(docx_path).resolve()
        pdf_path = Path(pdf_path).resolve()
        pdf_path.parent.mkdir(parents=True, exist_ok=True)

        if force_engine in ("auto", "univer"):
            if cls._try_univer_convert(docx_path, pdf_path):
                return True

        if force_engine in ("auto", "libreoffice"):
            return cls._try_libreoffice_convert(docx_path, pdf_path)

        return False

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
            return proc.returncode == 0 and pdf_path.exists()
        except Exception:
            return False

    @staticmethod
    def _try_libreoffice_convert(docx_path: Path, pdf_path: Path) -> bool:
        system = platform.system().lower()
        
        # On Windows, try Word/WPS COM first if available
        if system == "windows":
            try:
                import win32com.client
                word = win32com.client.Dispatch("Word.Application")
                doc = word.Documents.Open(str(docx_path), ReadOnly=1, Visible=0)
                doc.ExportAsFixedFormat(str(pdf_path), 17) # 17 = wdExportFormatPDF
                doc.Close(SaveChanges=0)
                word.Quit()
                if pdf_path.exists():
                    return True
            except Exception:
                pass

        # Cross-platform LibreOffice fallback
        soffice = shutil.which("soffice") or shutil.which("libreoffice")
        if not soffice:
            # Check common Windows install paths
            common_paths = [
                r"C:\Program Files\LibreOffice\program\soffice.exe",
                r"C:\Program Files (x86)\LibreOffice\program\soffice.exe"
            ]
            for cp in common_paths:
                if os.path.exists(cp):
                    soffice = cp
                    break

        if not soffice:
            return False

        try:
            outdir = str(pdf_path.parent)
            cmd = [soffice, "--headless", "--convert-to", "pdf", "--outdir", outdir, str(docx_path)]
            proc = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=60)
            
            # LibreOffice generates <stem>.pdf in outdir
            generated_pdf = pdf_path.parent / (docx_path.stem + ".pdf")
            if generated_pdf.exists() and generated_pdf != pdf_path:
                shutil.move(str(generated_pdf), str(pdf_path))

            return pdf_path.exists()
        except Exception:
            return False
