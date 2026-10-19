# MindReport 2.0 & CrossReport 使用手册与数据库调用指南

本文档详细说明了 MindReport 2.0 报表系统在 Windows / Linux 下的使用方式，包括 **独立可执行程序 (`MindReport.exe`)** 的使用说明、连接不同类型数据库的配置格式与端到端调用示例。

---

## 一、Windows 独立可执行程序 (`MindReport.exe`) 使用说明

### 1.1 程序概述
`dist\MindReport.exe` 是通过 Nuitka 编译构建的 **Windows x64 独立单文件可执行程序**，具备以下工业级特性：
* **零环境依赖**：运行目标机**无需安装 Python、pip 或任何 Python 第三方库**。
* **内置高性能 C# 引擎**：已内嵌 .NET 8 单文件数据引擎 `CrossReportRunner.exe`，自动实现异构数据库连接、毫秒级时序规整降采样与紧凑流式传输。
* **智能多级 PDF 路由**：按优先级 **`Office (Word/WPS) > LibreOffice > Chrome/Edge`** 自动检测与导出。在未安装任何 Office 的工控机或服务器上，能直接调用 Chrome / Edge 无头浏览器输出高清矢量 PDF。

### 1.2 命令行参数说明

```text
MindReport.exe [-h] [-p PRJPATH] -d DOCXTPL -s DSDEF [-a DSPARA] [-r VPARA] [-f PDFPATH]
```

| 参数 | 全称 | 必填 | 说明 |
| :--- | :--- | :---: | :--- |
| **`-d`** | `--docxtpl` | **是** | 待渲染的 Word Jinja 模板文件路径（`.docx`） |
| **`-s`** | `--dsdef` | **是** | 数据集定义 XML 文件路径（`.dsdef`） |
| **`-a`** | `--dspara` | 否 | 数据库连接参数字符串（格式详见第二章节） |
| **`-r`** | `--vpara` | 否 | 运行时变量宏替换参数（格式如 `BatchNo=batch_001`） |
| **`-f`** | `--pdfpath` | 否 | 导出的 PDF 文件路径（可选，未指定时仅生成 `*_autogen.docx`） |
| **`-p`** | `--prjpath` | 否 | 项目根目录路径（可选，用于相对路径定位） |

---

## 二、数据库连接参数格式规范 (`-a` / `--dspara`)

通过命令行参数 `-a`（或 `--dspara`）传入数据库连接字符串。多个属性之间使用分号 `;` 分隔，键值对使用等号 `=` 连接：

```text
DSName=<数据源名>;DBType=<类型>;ServerName=<主机>;Port=<端口>;DBName=<库名>;User=<用户>;Password=<密码>;DBPath=<文件路径>
```

> [!IMPORTANT]
> **数据源名称匹配 (`DSName`)**：
> `DSName` 的值必须与报表定义文件（`.dsdef`）中各 `<Table>` 节点的 `<DataSource>` 标签内容严格一致。例如在 `LYO.dsdef` 中数据源为 `LYO1`，则连接参数必须包含 `DSName=LYO1`。

### 2.1 支持的参数键（大小写不敏感）

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

### 2.2 数据库类型枚举对照表 (`DBType`)

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

## 三、各数据库完整调用示例

### 1. SQLite 数据库 (`DBType=4`)

#### 参数格式
```text
DSName=LYO1;DBType=4;DBPath=tests/data/lyo_test.db
```

#### 调用命令示例
- **使用单一可执行文件 `MindReport.exe`**：
  ```powershell
  .\dist\MindReport.exe `
    -d "assets/BatchManagerDemo/Report/LYO.docx" `
    -s "assets/BatchManagerDemo/Report/LYO.dsdef" `
    -a "DSName=LYO1;DBType=4;DBPath=tests/data/lyo_test.db" `
    -r "BatchNo=batch_001" `
    -f "output/LYO_Report.pdf"
  ```
- **使用 Python 源码运行**：
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
  bin\win-x64\CrossReportRunner.exe `
    -s "assets/BatchManagerDemo/Report/LYO.dsdef" `
    -a "DSName=LYO1;DBType=4;DBPath=tests/data/lyo_test.db" `
    -r "BatchNo=batch_001" `
    --export-json="output/lyo_out.json"
  ```

---

### 2. Microsoft SQL Server (`DBType=1`)

广泛用于工控 SCADA、InTouch、WinCC、Rockwell FactoryTalk。

#### 参数格式
```text
DSName=LYO1;DBType=1;ServerName=192.168.1.100;Port=1433;DBName=BatchManagerDB;User=sa;Password=YourPassword123
```

#### 调用命令示例
- **使用 `MindReport.exe`**：
  ```powershell
  .\dist\MindReport.exe `
    -d "assets/BatchManagerDemo/Report/LYO.docx" `
    -s "assets/BatchManagerDemo/Report/LYO.dsdef" `
    -a "DSName=LYO1;DBType=1;ServerName=192.168.1.100;Port=1433;DBName=BatchManagerDB;User=sa;Password=YourPassword123" `
    -r "BatchNo=batch_001" `
    -f "output/LYO_SQLServer_Report.pdf"
  ```

> [!NOTE]
> `CrossReportRunner` 内部默认启用了 `TrustServerCertificate=True;Timeout=30;`，原生支持自签名证书连接。

---

### 3. MySQL / MariaDB (`DBType=5`)

适用于开源技术栈 MES、IoT 采集网关。

#### 参数格式
```text
DSName=LYO1;DBType=5;ServerName=192.168.1.120;Port=3306;DBName=batch_tracking;User=root;Password=MySqlRootPassword!
```

#### 调用命令示例
- **使用 `MindReport.exe`**：
  ```powershell
  .\dist\MindReport.exe `
    -d "assets/BatchManagerDemo/Report/LYO.docx" `
    -s "assets/BatchManagerDemo/Report/LYO.dsdef" `
    -a "DSName=LYO1;DBType=5;ServerName=192.168.1.120;Port=3306;DBName=batch_tracking;User=root;Password=MySqlRootPassword!" `
    -r "BatchNo=batch_001" `
    -f "output/LYO_MySQL_Report.pdf"
  ```

---

### 4. PostgreSQL (`DBType=6`)

适用于半导体、生物制药高端装备与时序扩展数据库（TimescaleDB）。

#### 参数格式
```text
DSName=LYO1;DBType=6;Host=192.168.1.130;Port=5432;DBName=mes_production;User=postgres;Password=PgSecurePass2026
```

#### 调用命令示例
- **使用 `MindReport.exe`**：
  ```powershell
  .\dist\MindReport.exe `
    -d "assets/BatchManagerDemo/Report/LYO.docx" `
    -s "assets/BatchManagerDemo/Report/LYO.dsdef" `
    -a "DSName=LYO1;DBType=6;Host=192.168.1.130;Port=5432;DBName=mes_production;User=postgres;Password=PgSecurePass2026" `
    -r "BatchNo=batch_001" `
    -f "output/LYO_Postgres_Report.pdf"
  ```

---

### 5. Oracle Database (`DBType=2`)

适用于集团级 MES、历史数据仓库中心。

#### 参数格式（`DBName` 对应 Service Name 或 SID，如 `ORCL`）
```text
DSName=LYO1;DBType=2;ServerName=192.168.1.150;Port=1521;DBName=ORCL;User=c##lyo_report;Password=OracleReportPass123
```

#### 调用命令示例
- **使用 `MindReport.exe`**：
  ```powershell
  .\dist\MindReport.exe `
    -d "assets/BatchManagerDemo/Report/LYO.docx" `
    -s "assets/BatchManagerDemo/Report/LYO.dsdef" `
    -a "DSName=LYO1;DBType=2;ServerName=192.168.1.150;Port=1521;DBName=ORCL;User=c##lyo_report;Password=OracleReportPass123" `
    -r "BatchNo=batch_001" `
    -f "output/LYO_Oracle_Report.pdf"
  ```

---

### 6. Microsoft Access / OLE DB (`DBType=0` / `DBType=3`)

用于兼容遗留老旧系统的 `.mdb` / `.accdb` 文件。

#### 参数格式
```text
DSName=LYO1;DBType=0;DBPath=C:\LegacyData\MachineData.mdb
```

---

## 四、运行时实用技巧与环境配置

### 4.1 临时目录与磁盘空间注意事项
`MindReport.exe` 作为 Onefile 封装文件，在启动时会将内核与必要资源自解压至 Windows 的临时目录 `%TEMP%`（解压体积极小，运行结束后自动清理）。
* **如果操作系统 C 盘可用空间极低（小于 200MB）**，可能会引发 `[Error 112] There is not enough space on the disk`。
* **解决办法**：在运行命令前，将临时目录重定向至空间充裕的磁盘即可：
  * **CMD**：
    ```cmd
    set TEMP=D:\tmp&& set TMP=D:\tmp&& dist\MindReport.exe -d ...
    ```
  * **PowerShell**：
    ```powershell
    $env:TEMP="D:\tmp"; $env:TMP="D:\tmp"; .\dist\MindReport.exe -d ...
    ```

### 4.2 运行时宏替换参数 (`-r` / `--vpara`)
用于向 `.dsdef` 中定义的 SQL 查询注入动态参数（如批次号、起止时间等）：
```powershell
# 单个变量
-r "BatchNo=batch_001"

# 多个变量（分号分隔）
-r "BatchNo=batch_001;LineNo=L01;StartTime=2026-10-01 00:00:00;EndTime=2026-10-02 00:00:00"
```
SQL 中通过 `[%变量名%]` 引用：
```sql
SELECT * FROM LYO1_LYO_Header WHERE BatchNo = '[%BatchNo%]' AND LineNo = '[%LineNo%]'
```

### 4.3 仅生成数据 JSON 进行联调诊断
直接调用 `CrossReportRunner.exe` 进行 SQL 查询与时序规整验证：
```powershell
bin\win-x64\CrossReportRunner.exe `
  -s "assets/BatchManagerDemo/Report/LYO.dsdef" `
  -a "DSName=LYO1;DBType=4;DBPath=tests/data/lyo_test.db" `
  -r "BatchNo=batch_001" `
  --export-json="output/debug_data.json"
```

### 4.4 源码重新打包为 exe
如果修改了 Python 业务逻辑或模板解析规则，可在 Windows 开发机上运行预置脚本一键重新编译封包：
```powershell
.\scripts\build_nuitka.bat
```
输出的可执行文件将自动刷新到 [`dist\MindReport.exe`](file:///Z:/E/docxReport/dist/MindReport.exe)。
