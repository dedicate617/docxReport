# MindReport 2.0 & CrossReport 数据库连接调用手册

本文档详细说明了在 MindReport 2.0 报表系统与 `CrossReportRunner` 引擎中，针对不同数据库类型的连接配置格式与调用示例。

---

## 一、连接参数格式规范

通过命令行参数 `-a`（或 `--dspara`）传入数据库连接字符串。多个属性之间使用分号 `;` 分隔，键值对使用等号 `=` 连接：

```text
DSName=<数据源名>;DBType=<类型>;ServerName=<主机>;Port=<端口>;DBName=<库名>;User=<用户>;Password=<密码>;DBPath=<文件路径>
```

> [!IMPORTANT]
> **数据源名称匹配 (`DSName`)**：
> `DSName` 的值必须与报表定义文件（`.dsdef`）中各 `<Table>` 节点的 `<DataSource>` 标签内容严格一致。例如在 `LYO.dsdef` 中数据源为 `LYO1`，则连接参数必须包含 `DSName=LYO1`。

### 1.1 支持的参数键（大小写不敏感）

| 参数键名 | 别名支持 | 必填项 | 说明与默认值 |
| :--- | :--- | :---: | :--- |
| `DSName` | - | **是** | 数据源名称，需与 `.dsdef` 中的 `<DataSource>` 一致 |
| `DBType` | - | **是** | 数据库类型枚举值（数字 `0~6` 或字符串别名） |
| `ServerName` | `Host` | 视库而定 | 数据库服务器 IP 或域名（默认 `127.0.0.1`） |
| `Port` | - | 否 | 数据库端口号（未提供时按各数据库标准默认端口） |
| `DBName` | - | 视库而定 | 目标数据库名（Oracle 对应 Service Name 或 SID） |
| `User` | `UserID` | 视库而定 | 数据库登录用户名 |
| `Password` | `PassWord` | 视库而定 | 数据库登录密码 |
| `DBPath` | `Path` | SQLite必填 | 数据库文件的本地路径（SQLite / Access 专用） |

### 1.2 数据库类型枚举对照表 (`DBType`)

| 数字代码 | 别名字符串 | 对应数据库驱动 | 默认端口 | 跨平台支持 |
| :---: | :--- | :--- | :---: | :---: |
| **`4`** | `Sqlite` | `Microsoft.Data.Sqlite` | - | Windows / Linux / macOS |
| **`1`** | `MSSqlServer` | `Microsoft.Data.SqlClient` | 1433 | Windows / Linux / macOS |
| **`5`** | `MySql` | `MySqlConnector` | 3306 | Windows / Linux / macOS |
| **`6`** | `PostgreSql` | `Npgsql` | 5432 | Windows / Linux / macOS |
| **`2`** | `Oracle` | `Oracle.ManagedDataAccess.Core` | 1521 | Windows / Linux / macOS |
| **`0`** | `MSAccess` | OleDb Driver | - | 仅限 Windows（依赖本地 OLEDB 驱动） |
| **`3`** | `OleDB` | OleDb Driver | - | 仅限 Windows |

---

## 二、每种数据库的具体调用示例

### 1. SQLite 数据库 (`DBType=4` 或 `DBType=Sqlite`)
适用于嵌入式系统、边缘工业网关、单机版 MES/SCADA、本地离线仿真与单元测试。

#### 连接参数格式
```text
DSName=LYO1;DBType=4;DBPath=tests/data/lyo_test.db
```
或使用绝对路径：
```text
DSName=LYO1;DBType=Sqlite;DBPath=Z:\E\docxReport\tests\data\lyo_test.db
```

#### 调用命令示例
- **Python MindReport 完整流水线（Word + PDF）**：
  ```powershell
  python src/MindReport/main.py `
    -d "assets/BatchManagerDemo/Report/LYO.docx" `
    -s "assets/BatchManagerDemo/Report/LYO.dsdef" `
    -a "DSName=LYO1;DBType=4;DBPath=tests/data/lyo_test.db" `
    -r "BatchNo=batch_001" `
    -f "output/LYO_Report.pdf"
  ```
- **C# CrossReportRunner 原生调用**：
  ```powershell
  bin/win-x64/CrossReportRunner.exe `
    -s "assets/BatchManagerDemo/Report/LYO.dsdef" `
    -a "DSName=LYO1;DBType=4;DBPath=tests/data/lyo_test.db" `
    -r "BatchNo=batch_001" `
    --export-json="output/lyo_out.json"
  ```

---

### 2. Microsoft SQL Server (`DBType=1` 或 `DBType=MSSqlServer`)
广泛用于工业自动化 SCADA、AVEVA System Platform (InTouch/Wonderware)、西门子 WinCC、罗克韦尔 FactoryTalk 等系统。

#### 连接参数格式
- **标准默认端口（1433）**：
  ```text
  DSName=LYO1;DBType=1;ServerName=192.168.1.100;Port=1433;DBName=BatchManagerDB;User=sa;Password=YourPassword123
  ```
- **使用别名与自定义实例/端口**：
  ```text
  DSName=LYO1;DBType=MSSqlServer;Host=127.0.0.1;Port=14333;DBName=MES_MESDB;UserID=mes_report;PassWord=Report#Pass2026
  ```

#### 调用命令示例
```powershell
python src/MindReport/main.py `
  -d "assets/BatchManagerDemo/Report/LYO.docx" `
  -s "assets/BatchManagerDemo/Report/LYO.dsdef" `
  -a "DSName=LYO1;DBType=1;ServerName=192.168.1.100;Port=1433;DBName=BatchManagerDB;User=sa;Password=YourPassword123" `
  -r "BatchNo=batch_001" `
  -f "output/LYO_SQLServer_Report.pdf"
```

> [!NOTE]
> `CrossReportRunner` 内部默认启用了 `TrustServerCertificate=True;Timeout=30;`，支持自签名 SSL 证书连接。

---

### 3. MySQL / MariaDB (`DBType=5` 或 `DBType=MySql`)
适用于基于开源技术栈的 MES 制造执行系统、自研物联网平台与云边协同架构。

#### 连接参数格式
```text
DSName=LYO1;DBType=5;ServerName=192.168.1.120;Port=3306;DBName=batch_tracking;User=root;Password=MySqlRootPassword!
```

#### 调用命令示例
```powershell
python src/MindReport/main.py `
  -d "assets/BatchManagerDemo/Report/LYO.docx" `
  -s "assets/BatchManagerDemo/Report/LYO.dsdef" `
  -a "DSName=LYO1;DBType=5;ServerName=192.168.1.120;Port=3306;DBName=batch_tracking;User=root;Password=MySqlRootPassword!" `
  -r "BatchNo=batch_001" `
  -f "output/LYO_MySQL_Report.pdf"
```

---

### 4. PostgreSQL (`DBType=6` 或 `DBType=PostgreSql`)
适用于对数据完整性、时序扩展（TimescaleDB）要求极高的制药与半导体高端装备监控系统。

#### 连接参数格式
```text
DSName=LYO1;DBType=6;Host=192.168.1.130;Port=5432;DBName=mes_production;User=postgres;Password=PgSecurePass2026
```

#### 调用命令示例
```powershell
python src/MindReport/main.py `
  -d "assets/BatchManagerDemo/Report/LYO.docx" `
  -s "assets/BatchManagerDemo/Report/LYO.dsdef" `
  -a "DSName=LYO1;DBType=6;Host=192.168.1.130;Port=5432;DBName=mes_production;User=postgres;Password=PgSecurePass2026" `
  -r "BatchNo=batch_001" `
  -f "output/LYO_Postgres_Report.pdf"
```

---

### 5. Oracle Database (`DBType=2` 或 `DBType=Oracle`)
适用于大型制药企业集团级 ERP/MES 集成库、历史数据中心。

#### 连接参数格式
> 在 Oracle 中，`DBName` 对应的是数据库服务名（Service Name）或 SID（例如 `ORCL`、`FREE` 或 `HECL`）：

```text
DSName=LYO1;DBType=2;ServerName=192.168.1.150;Port=1521;DBName=ORCL;User=c##lyo_report;Password=OracleReportPass123
```

#### 调用命令示例
```powershell
python src/MindReport/main.py `
  -d "assets/BatchManagerDemo/Report/LYO.docx" `
  -s "assets/BatchManagerDemo/Report/LYO.dsdef" `
  -a "DSName=LYO1;DBType=2;ServerName=192.168.1.150;Port=1521;DBName=ORCL;User=c##lyo_report;Password=OracleReportPass123" `
  -r "BatchNo=batch_001" `
  -f "output/LYO_Oracle_Report.pdf"
```

---

### 6. Microsoft Access / OLE DB (`DBType=0` / `DBType=3`)
用于兼容遗留老旧系统的 `.mdb` / `.accdb` 文件。

#### 连接参数格式
```text
DSName=LYO1;DBType=0;DBPath=C:\LegacyData\MachineData.mdb
```

> [!WARNING]
> Access 和通用 OLEDB 依赖 Windows 本地驱动程序（如 ACE.OLEDB），在 Linux/容器环境中不支持。跨平台推荐优先迁移至 SQLite。

---

## 三、常用运行时参数与技巧

### 3.1 运行时宏替换参数 (`-r` / `--vpara`)
用于向 `.dsdef` 中定义的 SQL 查询注入动态参数（如批次号、起止时间、产线编号等）：

```powershell
# 单个变量
-r "BatchNo=batch_001"

# 多个变量（分号分隔）
-r "BatchNo=batch_001;LineNo=L01;StartTime=2026-10-01 00:00:00;EndTime=2026-10-02 00:00:00"
```

在 `.dsdef` 的 `<Sql>` 中通过 `[%变量名%]` 引用：
```sql
SELECT * FROM LYO1_LYO_Header WHERE BatchNo = '[%BatchNo%]' AND LineNo = '[%LineNo%]'
```

### 3.2 密码中包含特殊字符的处理
如果数据库连接密码中包含分号 `;` 或等号 `=`，建议在外部命令行使用单引号包裹整个 `-a` 字符串，或者转义传参。

### 3.3 仅生成数据 JSON 进行联调诊断
如果需要单独诊断 SQL 查询、时序数据 60 秒降采样是否正确，可以直接运行 `CrossReportRunner.exe` 输出 JSON：

```powershell
bin\win-x64\CrossReportRunner.exe `
  -s "assets/BatchManagerDemo/Report/LYO.dsdef" `
  -a "DSName=LYO1;DBType=4;DBPath=tests/data/lyo_test.db" `
  -r "BatchNo=batch_001" `
  --export-json="output/debug_data.json"
```
可在 `output/debug_data.json` 中直接检查查询出的全部表格行数据。
