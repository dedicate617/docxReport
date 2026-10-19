import argparse
import sys
import os
from pathlib import Path

repo_root = Path(__file__).resolve().parent.parent.parent
if str(repo_root) not in sys.path:
    sys.path.insert(0, str(repo_root))

from src.MindReport.runner_launcher import RunnerLauncher
from src.MindReport.context_builder import ContextBuilder
from src.MindReport.docx_renderer import DocxRenderer
from src.MindReport.pdf_router import PdfRouter

def run_report_pipeline(base_dir: Path, docxtpl: str, dsdef: str, dspara: str, vpara: str, pdfpath: str = "") -> int:
    try:
        # Step 1: Query database via CrossReportRunner stream pipeline
        print("[MindReport] Launching CrossReportRunner...")
        payload = RunnerLauncher.execute_and_stream(base_dir, dsdef, dspara, vpara)

        # Step 2: Build Jinja2 Context with LazyRowProxy
        context = ContextBuilder.build_template_context(payload)

        # Step 3: Render Word Template & update OpenXML Charts
        autogen_docx = Path(docxtpl).with_name(Path(docxtpl).stem + "_autogen.docx")
        print(f"[MindReport] Rendering Word template to {autogen_docx}...")
        DocxRenderer.render(docxtpl, context, str(autogen_docx), raw_payload=payload)
        print(f"[MindReport] Word report successfully generated: {autogen_docx}")

        # Step 4: Convert to PDF (optional, Univer first, LibreOffice fallback)
        if pdfpath:
            print(f"[MindReport] Converting to PDF {pdfpath}...")
            success = PdfRouter.convert_docx_to_pdf(autogen_docx, Path(pdfpath))
            if not success:
                print(f"[MindReport Notice] Current machine has no PDF rendering engine (Word/WPS/LibreOffice).")
                print(f"[MindReport Notice] The Word report is intact and fully populated at: {autogen_docx}")
                print(f"[MindReport Notice] To enable PDF export on headless/office-free machines, install LibreOffice:")
                print(f"  - Linux: sudo apt-get install libreoffice (or yum install libreoffice)")
                print(f"  - Windows: download LibreOffice Portable or install LibreOffice")
                return 2

        print("[MindReport] Report generation finished successfully.")
        return 0
    except Exception as ex:
        print(f"[MindReport Error] Pipeline failed: {ex}", file=sys.stderr)
        return 1

def main():
    parser = argparse.ArgumentParser(description="MindReport 2.0 Cross-Platform Runner")
    parser.add_argument("-p", "--prjpath", default="", help="Project path")
    parser.add_argument("-d", "--docxtpl", required=True, help="Word Jinja template path")
    parser.add_argument("-s", "--dsdef", required=True, help="Data query definition path")
    parser.add_argument("-a", "--dspara", default="", help="Database connection parameters")
    parser.add_argument("-r", "--vpara", default="", help="Runtime variable parameters")
    parser.add_argument("-f", "--pdfpath", default="", help="Output PDF destination path (optional)")

    args = parser.parse_args()
    base_dir = Path(__file__).resolve().parent.parent.parent

    code = run_report_pipeline(
        base_dir=base_dir,
        docxtpl=args.docxtpl,
        dsdef=args.dsdef,
        dspara=args.dspara,
        vpara=args.vpara,
        pdfpath=args.pdfpath
    )
    sys.exit(code)

if __name__ == "__main__":
    main()
