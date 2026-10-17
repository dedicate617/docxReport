import pytest
from pathlib import Path
from unittest.mock import patch
from src.MindReport.pdf_router import PdfRouter

def test_pdf_router_fallback_to_libreoffice_when_univer_fails(tmp_path):
    docx_file = tmp_path / "dummy.docx"
    docx_file.write_text("dummy")
    pdf_file = tmp_path / "dummy.pdf"

    # Mock univer failing and libreoffice succeeding
    with patch.object(PdfRouter, '_try_univer_convert', return_value=False), \
         patch.object(PdfRouter, '_try_libreoffice_convert', return_value=True):
        
        success = PdfRouter.convert_docx_to_pdf(docx_file, pdf_file)
        assert success is True
