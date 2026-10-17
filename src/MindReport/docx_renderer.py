from docxtpl import DocxTemplate
import jinja2
from datetime import datetime
from pathlib import Path
from src.MindReport.chart_updater import ChartUpdater

def datetimefmt(value, fmt='%Y-%m-%d %H:%M:%S'):
    if not value:
        return ""
    try:
        if isinstance(value, str):
            dt = datetime.fromisoformat(value.replace(' ', 'T'))
        elif isinstance(value, datetime):
            dt = value
        else:
            return str(value)
        return dt.strftime(fmt)
    except Exception:
        return str(value)

class DocxRenderer:
    @staticmethod
    def render(tpl_path: str, context: dict, output_path: str, raw_payload: dict = None):
        tpl = DocxTemplate(tpl_path)
        
        # Update embedded charts if raw payload with time series is provided
        if raw_payload and "Tables" in raw_payload:
            ChartUpdater.update_charts(tpl.get_docx(), raw_payload["Tables"])

        # Jinja2 environment with custom filters
        jinja_env = jinja2.Environment()
        jinja_env.filters['datetimefmt'] = datetimefmt

        tpl.render(context, jinja_env)
        
        out_p = Path(output_path)
        out_p.parent.mkdir(parents=True, exist_ok=True)
        tpl.save(str(out_p))
