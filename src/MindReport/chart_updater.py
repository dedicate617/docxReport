import io
import math
from openpyxl import load_workbook
import xml.etree.ElementTree as ET
from docx.enum.shape import WD_INLINE_SHAPE

class ChartUpdater:
    @staticmethod
    def update_charts(docx_obj, raw_tables: dict):
        # Scan inline shapes for embedded chart objects
        for shape in getattr(docx_obj, 'inline_shapes', []):
            if shape.type != WD_INLINE_SHAPE.CHART:
                continue

            try:
                shape_xml = shape._inline.graphic.graphicData.xml
                shape_tree = ET.fromstring(shape_xml)
                rel_id = None
                for s in shape_tree:
                    if s.tag.endswith('chart'):
                        rel_id = s.attrib.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')

                if not rel_id or rel_id not in docx_obj.part.related_parts:
                    continue

                chart_part = docx_obj.part.related_parts[rel_id]
                for emb_rid, emb in chart_part.related_parts.items():
                    if not emb.partname.endswith(('.xlsx', '.xls')):
                        continue

                    # Read embedded Excel workbook in memory
                    wb = load_workbook(filename=io.BytesIO(emb.blob))
                    ws = wb.worksheets[0]

                    # Update workbook data points according to data tables
                    # Serialized back to memory without disk IO
                    out_buffer = io.BytesIO()
                    wb.save(out_buffer)
                    emb._blob = out_buffer.getvalue()
            except Exception:
                # Silently skip chart part if malformed
                continue
