import argparse
import sys
import os
from pathlib import Path
from src.MindReport.runner_launcher import RunnerLauncher
from src.MindReport.context_builder import ContextBuilder
from src.MindReport.docx_renderer import DocxRenderer
from src.MindReport.pdf_router import PdfRouter

def run_report_pipeline(base_dir: Path, docxtpl: str, dsdef: str, dspara: str, vpara: str, pdfpath: str) -> int:
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

        # Step 4: Convert to PDF (Univer first, LibreOffice fallback)
        print(f"[MindReport] Converting to PDF {pdfpath}...")
        success = PdfRouter.convert_docx_to_pdf(autogen_docx, Path(pdfpath))
        if not success:
            print("[MindReport Error] All PDF conversion engines failed.")
            return 6

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
    parser.add_argument("-f", "--pdfpath", required=True, help="Output PDF destination path")

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
