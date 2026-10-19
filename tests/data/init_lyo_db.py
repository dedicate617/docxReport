import re
import sqlite3
import xml.etree.ElementTree as ET
from pathlib import Path

def init_db():
    data_dir = Path("tests/data")
    data_dir.mkdir(parents=True, exist_ok=True)
    db_path = data_dir / "lyo_test.db"
    if db_path.exists():
        db_path.unlink()

    dsdef_path = Path("assets/BatchManagerDemo/Report/LYO.dsdef")
    tree = ET.parse(dsdef_path)
    root = tree.getroot()

    conn = sqlite3.connect(str(db_path))

    for tbl in root.findall(".//Table"):
        name_elem = tbl.find("Name")
        if name_elem is None or not name_elem.text:
            continue
        table_name = name_elem.text.strip()
        sql_elem = tbl.find("Sql")
        sql = sql_elem.text.strip() if sql_elem is not None and sql_elem.text else ""

        # Extract all bracketed column names from SQL and from Fields
        cols = set()
        fields_elem = tbl.find("Fields")
        if fields_elem is not None:
            for f in fields_elem.findall("Field"):
                fn = f.find("Name")
                if fn is not None and fn.text:
                    cols.add(fn.text.strip())

        # Also regex all [ColName] in SQL
        bracket_cols = re.findall(r"\[([a-zA-Z0-9_]+)\]", sql)
        for bc in bracket_cols:
            if bc != table_name:
                cols.add(bc)

        # Always include BatchNo
        cols.add("BatchNo")

        cols_list = sorted(list(cols))
        col_defs = ", ".join([f"[{c}] TEXT" for c in cols_list])
        conn.execute(f"CREATE TABLE IF NOT EXISTS [{table_name}] ({col_defs})")

        # Insert realistic mock data for each table
        placeholders = ", ".join(["?" for _ in cols_list])

        if table_name == "LYO1_LYO_Header":
            row_data = {
                "ProjectNo": "PRJ_LYO_2026",
                "BatchNo": "batch_001",
                "BatchDescription": "冻干批次测试001",
                "BatchComment": "正常运行批次",
                "ProductName": "注射用冻干制剂",
                "RecipeName": "STANDARD_LYO_V2",
                "StartTime": "2026-10-17 08:00:00",
                "EndTime": "2026-10-18 10:30:00",
                "Model": "LYO-500",
                "InitiatorFullName": "张工"
            }
            conn.execute(f"INSERT INTO [{table_name}] VALUES ({placeholders})", [row_data.get(c, "") for c in cols_list])

        elif table_name == "LYO1_LYO_Recipe":
            recipes = [
                ("1", "预冻温度", "-45.0"),
                ("2", "预冻时间", "120"),
                ("3", "一次干燥真空", "0.08"),
                ("4", "一次干燥温度", "-15.0"),
                ("5", "解析干燥温度", "25.0")
            ]
            for r in recipes:
                rd = {"No": r[0], "Item": r[1], "Value": r[2], "BatchNo": "batch_001"}
                conn.execute(f"INSERT INTO [{table_name}] VALUES ({placeholders})", [rd.get(c, "") for c in cols_list])

        elif table_name == "LYO1_LYO_Parameter":
            params = [
                ("1", "板层温度", "30.0", "-50.0", "-44.8", "℃"),
                ("2", "冷阱温度", "-30.0", "-85.0", "-75.2", "℃"),
                ("3", "箱体压力", "100.0", "0.01", "0.065", "mbar"),
                ("4", "产品温度1", "25.0", "-45.0", "-42.1", "℃"),
            ]
            for p in params:
                rd = {"No": p[0], "Name": p[1], "HighLimit": p[2], "LowLimit": p[3], "Value": p[4], "Unit": p[5], "BatchNo": "batch_001"}
                conn.execute(f"INSERT INTO [{table_name}] VALUES ({placeholders})", [rd.get(c, "") for c in cols_list])

        elif table_name == "AlarmDataGroup":
            alarms = [
                ("ALM_001", "制冷异常", "制冷机高压报警", "警告", "2026-10-17 09:15:00", "2026-10-17 09:16:30", "已恢复", "batch_001"),
                ("ALM_002", "真空波动", "箱体压力偏高", "提示", "2026-10-17 14:20:00", "2026-10-17 14:22:10", "已恢复", "batch_001")
            ]
            for a in alarms:
                rd = {
                    "ID": a[0], "Name": a[1], "Message": a[2], "ALarmLevel": a[3],
                    "StartTime": a[4], "EndTime": a[5], "Status": a[6], "BatchNo": a[7],
                    "CurrentUser": "admin", "ForeColor": "0", "BackColor": "0", "ConfirmTime": a[5]
                }
                conn.execute(f"INSERT INTO [{table_name}] VALUES ({placeholders})", [rd.get(c, "") for c in cols_list])

        elif table_name == "LogDataGroup":
            logs = [
                ("2026-10-17 08:00:00", "批次启动", "admin", "操作员", "无", "batch_001"),
                ("2026-10-17 08:05:00", "配方下发确认", "admin", "复核员", "核对一致", "batch_001")
            ]
            for l in logs:
                rd = {"RecordTime": l[0], "Message": l[1], "Operater": l[2], "Verifier": l[3], "OperateComment": l[4], "BatchNo": l[5]}
                conn.execute(f"INSERT INTO [{table_name}] VALUES ({placeholders})", [rd.get(c, "") for c in cols_list])

        elif table_name == "LYO1_LYO_ProcessValue":
            # Generate 10 time series points for sampling verification
            base_temp = -45.0
            for i in range(15):
                t_str = f"2026-10-17 08:{i:02d}:00"
                rd = {
                    "BatchNo": "batch_001",
                    "RecordTime": t_str,
                    "RecordTimeMS": str(i * 1000),
                    "TE201": str(round(base_temp + i * 0.2, 2)),
                    "ShelfTemp": str(round(base_temp + i * 0.15, 2)),
                    "TE101": str(round(base_temp + i * 0.3, 2)),
                    "TE102": str(round(base_temp + i * 0.25, 2)),
                    "PT101": "0.065",
                    "PT301": "0.063",
                    "VacuumSetpoint": "0.08"
                }
                conn.execute(f"INSERT INTO [{table_name}] VALUES ({placeholders})", [rd.get(c, "") for c in cols_list])

        elif table_name == "LYO1_LYO_ProcessLog":
            logs = [
                ("2026-10-17 08:02:00", "进料完成，开始预冻阶段", "batch_001"),
                ("2026-10-17 08:10:00", "板层降温达到目标值 -45℃", "batch_001")
            ]
            for pl in logs:
                rd = {"RecordTime": pl[0], "Message": pl[1], "BatchNo": pl[2]}
                conn.execute(f"INSERT INTO [{table_name}] VALUES ({placeholders})", [rd.get(c, "") for c in cols_list])

    conn.commit()
    tables = [r[0] for r in conn.execute("SELECT name FROM sqlite_master WHERE type='table'").fetchall()]
    print(f"Initialized database {db_path} with {len(tables)} tables: {tables}")
    conn.close()

if __name__ == "__main__":
    init_db()
