import pytest
from src.MindReport.context_builder import ContextBuilder, LazyRowProxy

def test_lazy_row_proxy_attribute_and_item_access():
    columns = ["RecordTime", "TE101"]
    row_values = ["2026-10-17 12:00:00", 25.5]
    proxy = LazyRowProxy(columns, row_values)

    assert proxy.RecordTime == "2026-10-17 12:00:00"
    assert proxy["TE101"] == 25.5
    assert proxy.get("NonExistent", "default") == "default"

def test_context_builder_expands_tables_as_list_of_proxies():
    payload = {
        "Variables": {"BatchNo": "B001"},
        "Tables": {
            "LYO1_ProcessValue": {
                "columns": ["RecordTime", "TE101"],
                "rows": [
                    ["2026-10-17 12:00:00", 25.5],
                    ["2026-10-17 12:01:00", 25.6]
                ]
            }
        }
    }
    context = ContextBuilder.build_template_context(payload)
    assert context["Variables"]["BatchNo"] == "B001"
    assert len(context["LYO1_ProcessValue"]) == 2
    assert context["LYO1_ProcessValue"][0].TE101 == 25.5
