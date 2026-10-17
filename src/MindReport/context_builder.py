from typing import List, Any, Dict

class LazyRowProxy:
    __slots__ = ('_columns', '_values', '_map')

    def __init__(self, columns: List[str], values: List[Any]):
        self._columns = columns
        self._values = values
        self._map = {col.lower(): val for col, val in zip(columns, values)}

    def __getattr__(self, name: str) -> Any:
        try:
            return self._map[name.lower()]
        except KeyError:
            raise AttributeError(f"'LazyRowProxy' has no attribute '{name}'")

    def __getitem__(self, item: str) -> Any:
        return self._map[item.lower()]

    def get(self, key: str, default: Any = None) -> Any:
        return self._map.get(key.lower(), default)

    def to_dict(self) -> Dict[str, Any]:
        return {col: val for col, val in zip(self._columns, self._values)}

class ContextBuilder:
    @staticmethod
    def build_template_context(payload: dict) -> dict:
        context = {
            "Variables": payload.get("Variables") or payload.get("variables") or {}
        }
        tables = payload.get("Tables") or payload.get("tables") or {}
        for table_name, table_data in tables.items():
            cols = table_data.get("Columns") or table_data.get("columns") or []
            rows = table_data.get("Rows") or table_data.get("rows") or []
            context[table_name] = [LazyRowProxy(cols, r) for r in rows]
        return context
