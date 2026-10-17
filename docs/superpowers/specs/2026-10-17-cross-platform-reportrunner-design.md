# 跨平台报表运行器与关联调度器 (CrossReportRunner & MindReport 2.0) 技术设计规范

## 1. 背景与目标

### 1.1 现状与挑战
现有报表系统基于 Windows 平台的 .NET Framework 4.0 和 Python 2.7 构建：
* **数据查询端**：`MESReportRunnerShell.exe` 是基于 WinForms 的 .NET 4.0 程序，仅能在 Windows 上运行，依赖宿主界面消息循环，通过落盘导出 `xml` / `xsd` 文件。
* **渲染与导出端**：`MindReport.py` 基于 Python 2.7，通过子进程调用 `MESReportRunnerShell.exe`，解析 XML 并转换为 JSON，利用 `docxtpl` 渲染 Word 模板并操纵 Word 内置 OpenXML 图表，最后通过 Windows COM 接口（Office/WPS）或外部 LibreOffice 转换为 PDF。
* **业务痛点**：
  * 无法在 Linux x86 及 Linux ARM（如工控机 ARM32/armv7l、ARM64/aarch64）边缘设备上运行；
  * 大批量时序数据（如冻干机连续数天的过程变量）经由磁盘 XML/JSON 文件中转，在 ARM 工控机低速闪存（eMMC/SD卡）上产生巨大的磁盘 I/O 损耗与卡顿；
  * Python 2.7 已废弃，缺乏现代异步与类型注解支持。

### 1.2 核心目标
1. **多平台无依赖运行**：在 `Windows x86`、`Windows x64`、`Linux x86_64`、`Linux ARM32 (armv7l)`、`Linux ARM64 (aarch64)` 上完全独立运行，无需在目标系统安装 .NET SDK/Runtime。
2. **多数据源纯托管支持**：全面支持 SQL Server、SQLite、MySQL、PostgreSQL、Oracle 5 大主流数据库，全部采用 100% C# 纯托管驱动，无原生 C/C++ 动态链接库依赖。
3. **兼容原报表设计器**：继续使用原 Windows 报表设计器配置数据源与查询结构，直接消费其导出的 `.dsdef` 数据集定义文件。
4. **低内存与零磁盘 I/O 通道**：重写进程间数据交换，采用紧凑矩阵数据模型与标准输出 Gzip 内存流式管道传输，消除中间磁盘写入，并将内存峰值严格控制在几十兆以内。
5. **Jinja2 模板渲染与图表自动绑定**：完整继承 Word Jinja2 语法支持，自动根据抽样数据重写 Word 内置 OpenXML 图表及嵌入的 Excel 数据集。
6. **双 PDF 转换引擎**：首选尝试 `univer.ai` 无头渲染组件，若环境未配置或转换异常，自动无缝回退至 `LibreOffice Headless`。
7. **自动化 CI/CD 编译**：提供在线交叉编译流水线，一键产出全架构发布包。

---

## 2. 系统总体架构与数据流管道

系统采用**双层解耦、协同流式驱动**的架构：

```mermaid
flowchart TD
    subgraph Design["1. 设计阶段 (Windows 开发机)"]
        Designer["MES 报表设计器"] -->|导出| DSDEF["*.dsdef (数据集定义)"]
        Word["Word / WPS"] -->|编写 Jinja2 语法| DOCX["*.docx (报表模板)"]
    end

    subgraph Query["2. 数据查询阶段 (CrossReportRunner - .NET 8)"]
        CLI_In["CLI 参数: --dsdef, --dspara, --vpara"] --> Runner["CrossReportRunner 独立二进制"]
        DSDEF --> Runner
        DB[("5类数据库: SQLServer / SQLite\nMySQL / Postgres / Oracle")] <-->|纯托管 ADO.NET 驱动| Runner
        Runner -->|stdout Gzip 内存流式管道 (Zero Disk IO)| Pipe["内存压缩流 (stdout)"]
    end

    subgraph Render["3. 渲染与导出阶段 (MindReport 2.0 - Python 3)"]
        Pipe --> Launcher["Runner 启动器与流式解压"]
        DOCX --> Engine["DocxRenderer (docxtpl)"]
        Launcher --> Builder["ContextBuilder (LazyRowProxy)"]
        Builder --> Engine
        Builder --> Chart["OpenXML ChartUpdater (折线图/柱状图)"]
        Chart --> GenDocx["*_autogen.docx"]
        Engine --> GenDocx
        
        GenDocx --> Router{"PDF 转换路由适配器"}
        Router -->|优先尝试| Univer["Univer.ai Headless Engine (Node.js)"]
        Router -->|失败/环境受限降级| LibreOffice["LibreOffice Headless / OS COM"]
        Univer --> OutputPDF["最终交付: *.pdf"]
        LibreOffice --> OutputPDF
    end
```

---

## 3. 跨平台数据查询运行器 (CrossReportRunner) 设计

### 3.1 运行时与多架构发布
* **技术选型**：.NET 8 (C# 12)
* **发布配置**：
  * `PublishSingleFile=true`（单文件自包含二进制）
  * `SelfContained=true`（内置 .NET 运行时，目标机器免装环境）
  * `IncludeNativeLibrariesForSelfExtract=true`
  * `EnableCompressionInSingleFile=true`
* **目标架构 RID 矩阵**：
  * `win-x86`
  * `win-x64`
  * `linux-x64`
  * `linux-arm`（支持 ARM32v7）
  * `linux-arm64`（支持 ARM64v8）

### 3.2 纯托管数据库驱动矩阵 (Zero Native Dependency)
所有驱动均使用 100% 托管 C# 实现，彻底避免 Linux ARM32 缺少动态链接库的问题：
* **Microsoft SQL Server**: `Microsoft.Data.SqlClient` (v5.2+)
* **SQLite**: `Microsoft.Data.Sqlite` (v8.0+)
* **MySQL / MariaDB**: `MySqlConnector` (v2.3+)
* **PostgreSQL**: `Npgsql` (v8.0+)
* **Oracle**: `Oracle.ManagedDataAccess.Core` (v23.4+)

### 3.3 核心查询与抽样引擎模块
1. **`DsDefParser`**：
   * 解析 `.dsdef` XML 文档，提取 `<Table>` 节点、`<Sql>` 查询语句、`<EnableSampling>` 抽样配置、`<Union>` 联合表配置、`<RelateFields>` 关联字段及字段数据类型定义。
2. **`SqlParameterResolver`**：
   * **运行时变量替换**：将 SQL 中形如 `[%BatchNo%]` 的占位符替换为传入的实际变量值（忽略大小写）；
   * **父子嵌套关联**：对层级子查询，提取父表当前行的字段值替换子表 SQL 中的 `[#ParentField#]` 占位符。
3. **`SamplingEngine`**：
   * **时间节拍抽样（Regularize = 0）**：根据 `TimeUnit`（0:Year, 1:Month, 2:Day, 3:Hour, 4:Minute, 5:Second）和 `SamplingDistance`，对 `SamplingField`（如 `RecordTime`）划分等距时间窗口，每窗口抽取基准值。
   * **数据节拍抽样（Drift = 1）**：根据设定偏差阈值过滤稳定平稳数据，保留突变与拐点数据。
4. **`UnionAndRelateManager`**：
   * 支持指定 `<Union>` 目标表的数据合并（如 `LYO1_LYO_ProcessLog` 合并到 `LYO1_LYO_ProcessValue`）；
   * 依据 `<SecondarySort>`（如 `RecordTime ASC`）执行全表重新排序。

### 3.4 低内存与零 I/O 传输管道
* **命令行开关**：
  * `--output-mode=pipe-gzip`（默认）：输出直接流向标准输出。
  * `--output-mode=file`：输出到物理磁盘文件（仅在指定 `--export-json` 或 `--export-xml` 时用于离线排查）。
* **流式管道写入器 (`StreamSerializer`)**：
  * 使用 `System.Text.Json.Utf8JsonWriter` 直接包装 `GZipStream(Console.OpenStandardOutput(), CompressionLevel.Fastest)`；
  * 不在托管堆中构建大字符串，行数据边读边压边推入操作系统管道；
  * 采用**紧凑列式结构**：
    ```json
    {
      "TableName": {
        "columns": ["Col1", "Col2", "..."],
        "rows": [
          ["val1", 123, 45.6],
          ["val2", 789, 12.3]
        ]
      },
      "Variables": {
        "BatchNo": "batch_001",
        "CurrentUser": "operator"
      }
    }
    ```

---

## 4. 报表渲染与双 PDF 转换调度器 (MindReport 2.0) 设计

### 4.1 核心架构与模块划分
基于 **Python 3.10+** 构建，采用模块化设计：
* `main.py`：CLI 入口，使用 `argparse` 处理命令行参数，100% 兼容历史参数；
* `runner_launcher.py`：根据当前 OS 与 CPU 架构自动寻路并启动对应的 `CrossReportRunner` 二进制；通过管道流式接收 Gzip 数据并反序列化；
* `context_builder.py`：构建 Jinja2 模板上下文，提供 `LazyRowProxy` 兼顾模板字典访问体验与低内存开销；
* `docx_renderer.py`：基于 `docxtpl.DocxTemplate` 进行模板渲染，注册 `datetimefmt` 等自定义过滤器；
* `chart_updater.py`：解析 Word OpenXML 图表（`WD_INLINE_SHAPE.CHART`），更新嵌入的 Excel 数据集与图表 XML 缓存；
* `pdf_router.py`：双引擎转换路由器，优先尝试 Univer，回退执行 LibreOffice。

### 4.2 Word OpenXML 图表动态更新机制
针对 `LYO.docx` 中的工艺曲线图表：
1. 遍历 Word 内联形状，定位类型为 `CHART` 的部件及其关联关系（`r:id`）；
2. 提取绑定的嵌入式 Excel 工作簿（`emb.blob`），通过 `openpyxl` 在内存中覆写时序类别与数值列；
3. 更新图表定义 XML 中的 `<c:cat>`（分类轴缓存 `<c:strCache>`）与 `<c:val>`（数据点缓存 `<c:numCache>`）；
4. 支持根据横轴密集度配置主刻度间隔（`segment` 步长），确保曲线图坐标轴刻度整洁。

### 4.3 弹性双 PDF 转换引擎路由
```text
                  [待转换 *_autogen.docx]
                             │
                             ▼
              ┌─────────────────────────────┐
              │ 1. 尝试 Univer.ai 无头适配器 │
              └──────────────┬──────────────┘
                             │
                  ┌──────────┴──────────┐
                  │ 是否成功产出 PDF?    │
                  ├─────────────────────┤
            [成功]│                     │[失败/未部署/ARM32受限]
                  ▼                     ▼
          ┌──────────────┐    ┌──────────────────────────┐
          │ 返回成功 (0)  │    │ 2. 自动降级至方案 1:      │
          └──────────────┘    │ Linux: LibreOffice       │
                              │ Windows: Word/WPS/Libre  │
                              └────────────┬─────────────┘
                                           │
                                           ▼
                                  ┌──────────────────┐
                                  │   返回最终 PDF   │
                                  └──────────────────┘
```
1. **Univer.ai 无头适配器 (`univer-convert.js`)**：
   * 基于 Node.js 和 Univer Facade / Exchange API；
   * 调用命令：`node src/univer_adapter/univer-convert.js <input.docx> <output.pdf>`；
   * 设置超时熔断机制（默认 30 秒）。
2. **LibreOffice / OS 自动化保底**：
   * Linux：`soffice --headless --convert-to pdf --outdir <dir> <input.docx>`；
   * Windows：优先尝试 MS Word COM 接口，次选 WPS COM 接口，保底使用 LibreOffice。

---

## 5. CI/CD 自动化线上编译流水线

支持 GitHub Actions 与私有 GitLab CI，单 Runner 交叉编译全平台发布件：

### 5.1 自动化构建流程
1. **编译矩阵 (`matrix`)**：在单台 `ubuntu-latest` 容器中利用 .NET 8 SDK 并发交叉编译出 5 种独立二进制；
2. **自动化测试**：拉起 Linux-x64 运行器，使用真实 `LYO.dsdef` 和 SQLite/Mock 数据库进行全链路端到端渲染及 PDF 转换冒烟测试；
3. **发布打包 (`release`)**：将各架构二进制文件与 Python 调度源码归档为 `MindReport-Release-v*.tar.gz` 自动化发布。

---

## 6. 错误处理与诊断体系

### 6.1 退出码规范
* `0`: 成功完成报表生成与 PDF 输出；
* `1`: 命令行参数解析失败或缺少必填项；
* `2`: `.dsdef` 文件不存在或 XML 格式不合法；
* `3`: 数据库连接握手失败（超时/鉴权错误）；
* `4`: SQL 执行异常或抽样字段不存在；
* `5`: Word 模板 Jinja2 渲染或 OpenXML 图表解析错误；
* `6`: PDF 转换双引擎均无法产出有效文件。

### 6.2 安全脱敏与日志审计
* 日志文件统一输出至 `md.log`（并保留控制台标准错误输出）；
* 自动屏蔽 `--dspara` 连接字符串中的敏感字段（如 `Password=******`、`User Id=***`），防止生产环境密码泄露。

---

## 7. 验证与验收标准

1. **功能验收**：
   * 使用原报表设计器导出的 `assets/BatchManagerDemo/Report/LYO.dsdef`，能够正确在 SQLite、SQL Server 等数据源下查询出数据；
   * 抽样后输出的数据集行数与字段值与现存 `LYO.xml` / `LYO.json` 完全一致；
   * 生成的 `LYO.pdf` 包含完整的工艺曲线图表、配方表格与告警记录。
2. **性能与内存验收**：
   * 在 Linux ARM32 / ARM64 环境下，处理 50,000+ 条时序数据时，进程峰值内存占用不超过 80MB；
   * 数据交换阶段不产生任何临时中间数据文件（零磁盘 I/O 写入）。
3. **跨平台兼容验收**：
   * 独立二进制在干净的 Linux 系统（无 .NET 运行环境）上无需任何安装依赖，直接赋予执行权限即可成功运行。
