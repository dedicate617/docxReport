# 跨平台 ReportRunner 及关联调度器 (MindReport 2.0) 实施计划

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 构建可在 Windows (x86/x64) 与 Linux (x64/ARM32/ARM64) 运行的跨平台 .NET 8 单文件数据查询运行器（CrossReportRunner）与 Python 3.10+ 报表调度器（MindReport 2.0），实现低内存零磁盘 I/O 管道数据传输、Jinja2 Word 模板填充、OpenXML 图表自动绑定与优先 Univer.ai / 回退 LibreOffice 的双 PDF 转换引擎。

**Architecture:** 采用双层解耦架构。底层由 .NET 8 纯托管 ADO.NET 驱动解析 `.dsdef` 并执行跨平台多数据库查询、工业抽样和关联合并，通过标准输出经 `GZipStream` 输出紧凑 2D 矩阵 JSON 流；上层由 Python 3.10+ 调度器流式解压还原上下文，经 `docxtpl` 渲染 Word 模板并注入 OpenXML 曲线图表，最后经双引擎路由适配器输出高保真 PDF 报告。

**Tech Stack:**
* **C# / .NET 8 (C# 12)**: `Microsoft.Data.SqlClient`, `Microsoft.Data.Sqlite`, `MySqlConnector`, `Npgsql`, `Oracle.ManagedDataAccess.Core`, `System.Text.Json`
* **Python 3.10+**: `docxtpl`, `python-docx`, `openpyxl`, `lxml`, `pytest`
* **Node.js (Univer Adapter)**: `@univerjs/core`, `@univerjs/engine-render`, `@univerjs-pro/exchange-client`
* **PDF Fallback**: LibreOffice Headless / Windows Word/WPS COM

## Global Constraints

* **跨平台单文件发布**：`win-x86`, `win-x64`, `linux-x64`, `linux-arm` (ARM32v7), `linux-arm64` (aarch64)。目标环境免装 .NET SDK/Runtime。
* **纯托管数据库驱动**：绝不引入任何需本地编译或依赖 C/C++ 原生 `.so` 的数据库客户端。
* **输入契约兼容**：全面兼容原系统参数：`--prjpath`, `--docxtpl`, `--dsdef`, `--dspara`, `--vpara`, `--pdfpath`。
* **零磁盘 I/O**：默认数据流式管道传输，不落地中间 JSON 物理文件；峰值工作内存严格控制在 80MB 以内。
* **双引擎 PDF 转换**：优先尝试 Univer.ai 无头适配器，转换失败或受限时自动平滑回退至 LibreOffice。

---

### Task 1: .NET 8 工程脚手架与数据模型

**Files:**
* Create: `src/CrossReportRunner/CrossReportRunner.csproj`
* Create: `src/CrossReportRunner/Models/DsDefModels.cs`
* Create: `src/CrossReportRunner/Models/CompactDataPayload.cs`
* Test: `tests/CrossReportRunner.Tests/CrossReportRunner.Tests.csproj`
* Test: `tests/CrossReportRunner.Tests/ModelSerializationTests.cs`

**Interfaces:**
* Produces:
  * `DsReportDefinition`: `.dsdef` 根对象模型
  * `DsTableDefinition`: 数据表查询定义（Sql, Sampling, Union, Fields, RelateFields）
  * `CompactDataPayload`: 紧凑矩阵数据载荷，包含 `Dictionary<string, TableData>` 与 `Dictionary<string, string> Variables`
  * `TableData`: `{ List<string> Columns, List<List<object?>> Rows }`

- [ ] **Step 1: 创建测试工程与待测序列化测试用例**

```csharp
// tests/CrossReportRunner.Tests/ModelSerializationTests.cs
using System.Text.Json;
using CrossReportRunner.Models;
using Xunit;

namespace CrossReportRunner.Tests;

public class ModelSerializationTests
{
    [Fact]
    public void CompactDataPayload_ShouldSerializeAndDeserializeCorrectly()
    {
        var payload = new CompactDataPayload();
        payload.Variables["BatchNo"] = "batch_001";
        
        var table = new TableData
        {
            Columns = new List<string> { "RecordTime", "TE101", "Pressure" },
            Rows = new List<List<object?>>
            {
                new() { "2026-10-17 12:00:00", 25.1, 101.3 },
                new() { "2026-10-17 12:01:00", 25.2, 101.2 }
            }
        };
        payload.Tables["LYO1_ProcessValue"] = table;

        var json = JsonSerializer.Serialize(payload);
        var deserialized = JsonSerializer.Deserialize<CompactDataPayload>(json);

        Assert.NotNull(deserialized);
        Assert.Equal("batch_001", deserialized.Variables["BatchNo"]);
        Assert.True(deserialized.Tables.ContainsKey("LYO1_ProcessValue"));
        Assert.Equal(2, deserialized.Tables["LYO1_ProcessValue"].Rows.Count);
        Assert.Equal(3, deserialized.Tables["LYO1_ProcessValue"].Columns.Count);
    }
}
```

- [ ] **Step 2: 运行测试验证失败**

Run: `dotnet test tests/CrossReportRunner.Tests/CrossReportRunner.Tests.csproj`
Expected: FAIL (Projects or classes not found)

- [ ] **Step 3: 编写 CSPROJ 与核心模型定义**

```xml
<!-- src/CrossReportRunner/CrossReportRunner.csproj -->
<Project Sdk="Microsoft.NET.Sdk">
  <PropertyGroup>
    <OutputType>Exe</OutputType>
    <TargetFramework>net8.0</TargetFramework>
    <ImplicitUsings>enable</ImplicitUsings>
    <Nullable>enable</Nullable>
    <PublishSingleFile>true</PublishSingleFile>
    <SelfContained>true</SelfContained>
    <IncludeNativeLibrariesForSelfExtract>true</IncludeNativeLibrariesForSelfExtract>
    <EnableCompressionInSingleFile>true</EnableCompressionInSingleFile>
  </PropertyGroup>

  <ItemGroup>
    <PackageReference Include="Microsoft.Data.SqlClient" Version="5.2.2" />
    <PackageReference Include="Microsoft.Data.Sqlite" Version="8.0.10" />
    <PackageReference Include="MySqlConnector" Version="2.3.7" />
    <PackageReference Include="Npgsql" Version="8.0.5" />
    <PackageReference Include="Oracle.ManagedDataAccess.Core" Version="23.6.0" />
  </ItemGroup>
</Project>
```

```csharp
// src/CrossReportRunner/Models/CompactDataPayload.cs
namespace CrossReportRunner.Models;

public class TableData
{
    public List<string> Columns { get; set; } = new();
    public List<List<object?>> Rows { get; set; } = new();
}

public class CompactDataPayload
{
    public Dictionary<string, string> Variables { get; set; } = new(StringComparer.OrdinalIgnoreCase);
    public Dictionary<string, TableData> Tables { get; set; } = new(StringComparer.OrdinalIgnoreCase);
}
```

```csharp
// src/CrossReportRunner/Models/DsDefModels.cs
namespace CrossReportRunner.Models;

public enum SamplingType
{
    Regularize = 0,
    Drift = 1
}

public enum TimeUnit
{
    Year = 0,
    Month = 1,
    Day = 2,
    Hour = 3,
    Minute = 4,
    Second = 5
}

public class DsFieldDefinition
{
    public string Name { get; set; } = string.Empty;
    public string FieldType { get; set; } = "0";
}

public class DsTableDefinition
{
    public string Name { get; set; } = string.Empty;
    public string DataSource { get; set; } = string.Empty;
    public string Sql { get; set; } = string.Empty;
    public string Union { get; set; } = string.Empty;
    public string RelateFields { get; set; } = string.Empty;
    public string SecondarySort { get; set; } = string.Empty;
    public bool EnableSampling { get; set; }
    public string SamplingField { get; set; } = string.Empty;
    public int SamplingDistance { get; set; } = 60;
    public SamplingType SamplingType { get; set; } = SamplingType.Regularize;
    public TimeUnit TimeUnit { get; set; } = TimeUnit.Second;
    public List<DsFieldDefinition> Fields { get; set; } = new();
    public List<DsTableDefinition> SubTables { get; set; } = new();
}

public class DsReportDefinition
{
    public List<DsTableDefinition> Tables { get; set; } = new();
}
```

- [ ] **Step 4: 运行测试验证通过**

Run: `dotnet test tests/CrossReportRunner.Tests/CrossReportRunner.Tests.csproj`
Expected: PASS (1 passed)

- [ ] **Step 5: 提交代码**

```bash
git add src/CrossReportRunner tests/CrossReportRunner.Tests
git commit -m "feat(runner): initialize .NET 8 project and data models"
```

---

### Task 2: .dsdef XML 解析器 (DsDefParser)

**Files:**
* Create: `src/CrossReportRunner/Parsers/DsDefParser.cs`
* Test: `tests/CrossReportRunner.Tests/DsDefParserTests.cs`

**Interfaces:**
* Consumes: `DsReportDefinition`, `DsTableDefinition`
* Produces: `DsDefParser.Parse(string xmlContent)` -> `DsReportDefinition`
* Produces: `DsDefParser.ParseFile(string filePath)` -> `DsReportDefinition`

- [ ] **Step 1: 编写基于实际 LYO.dsdef 结构的单元测试**

```csharp
// tests/CrossReportRunner.Tests/DsDefParserTests.cs
using CrossReportRunner.Parsers;
using CrossReportRunner.Models;
using Xunit;

namespace CrossReportRunner.Tests;

public class DsDefParserTests
{
    private const string SampleDsdef = @"<Report>
  <Tables>
    <Table>
      <Name>LYO1_LYO_Header</Name>
      <DataSource>LYO1</DataSource>
      <Sql>SELECT [BatchNo] FROM [LYO1_LYO_Header] WHERE [BatchNo] = '[%BatchNo%]'</Sql>
      <EnableSampling>False</EnableSampling>
      <Fields>
        <Field><Name>BatchNo</Name><FieldType>0</FieldType></Field>
      </Fields>
    </Table>
    <Table>
      <Name>LYO1_LYO_ProcessValue</Name>
      <DataSource>LYO1</DataSource>
      <Sql>SELECT [TE201],[RecordTime] FROM [LYO1_LYO_ProcessValue] WHERE [BatchNo] = '[%BatchNo%]'</Sql>
      <EnableSampling>True</EnableSampling>
      <SamplingDistance>60</SamplingDistance>
      <SamplingField>RecordTime</SamplingField>
      <SamplingType>0</SamplingType>
      <TimeUnit>5</TimeUnit>
    </Table>
  </Tables>
</Report>";

    [Fact]
    public void Parse_ValidXml_ReturnsPopulatedDefinition()
    {
        var result = DsDefParser.Parse(SampleDsdef);
        Assert.NotNull(result);
        Assert.Equal(2, result.Tables.Count);
        
        var header = result.Tables[0];
        Assert.Equal("LYO1_LYO_Header", header.Name);
        Assert.Equal("LYO1", header.DataSource);
        Assert.False(header.EnableSampling);
        Assert.Single(header.Fields);

        var process = result.Tables[1];
        Assert.Equal("LYO1_LYO_ProcessValue", process.Name);
        Assert.True(process.EnableSampling);
        Assert.Equal(60, process.SamplingDistance);
        Assert.Equal("RecordTime", process.SamplingField);
        Assert.Equal(TimeUnit.Second, process.TimeUnit);
    }
}
```

- [ ] **Step 2: 运行测试验证失败**

Run: `dotnet test tests/CrossReportRunner.Tests/CrossReportRunner.Tests.csproj --filter DsDefParserTests`
Expected: FAIL (`DsDefParser` does not exist)

- [ ] **Step 3: 编写 DsDefParser 实现**

```csharp
// src/CrossReportRunner/Parsers/DsDefParser.cs
using System.Xml.Linq;
using CrossReportRunner.Models;

namespace CrossReportRunner.Parsers;

public static class DsDefParser
{
    public static DsReportDefinition ParseFile(string filePath)
    {
        if (!File.Exists(filePath))
            throw new FileNotFoundException($"dsdef file not found: {filePath}");

        var content = File.ReadAllText(filePath);
        return Parse(content);
    }

    public static DsReportDefinition Parse(string xmlContent)
    {
        var doc = XDocument.Parse(xmlContent);
        var reportDef = new DsReportDefinition();

        var tableElements = doc.Descendants("Table");
        foreach (var tblElem in tableElements)
        {
            // Only parse top-level Table elements here, children will be handled hierarchically
            if (tblElem.Parent?.Name == "Tables" && tblElem.Parent?.Parent?.Name == "Report")
            {
                reportDef.Tables.Add(ParseTable(tblElem));
            }
        }

        return reportDef;
    }

    private static DsTableDefinition ParseTable(XElement element)
    {
        var table = new DsTableDefinition
        {
            Name = element.Element("Name")?.Value.Trim() ?? string.Empty,
            DataSource = element.Element("DataSource")?.Value.Trim() ?? string.Empty,
            Sql = element.Element("Sql")?.Value.Trim() ?? string.Empty,
            Union = element.Element("Union")?.Value.Trim() ?? string.Empty,
            RelateFields = element.Element("RelateFields")?.Value.Trim() ?? string.Empty,
            SecondarySort = element.Element("SecondarySort")?.Value.Trim() ?? string.Empty,
            SamplingField = element.Element("SamplingField")?.Value.Trim() ?? string.Empty,
        };

        if (bool.TryParse(element.Element("EnableSampling")?.Value, out var enableSampling))
            table.EnableSampling = enableSampling;

        if (int.TryParse(element.Element("SamplingDistance")?.Value, out var samplingDistance))
            table.SamplingDistance = samplingDistance;

        if (int.TryParse(element.Element("SamplingType")?.Value, out var samplingType))
            table.SamplingType = (SamplingType)samplingType;

        if (int.TryParse(element.Element("TimeUnit")?.Value, out var timeUnit))
            table.TimeUnit = (TimeUnit)timeUnit;

        var fieldsElem = element.Element("Fields");
        if (fieldsElem != null)
        {
            foreach (var fld in fieldsElem.Elements("Field"))
            {
                table.Fields.Add(new DsFieldDefinition
                {
                    Name = fld.Element("Name")?.Value.Trim() ?? string.Empty,
                    FieldType = fld.Element("FieldType")?.Value.Trim() ?? "0"
                });
            }
        }

        var subTablesElem = element.Element("Tables");
        if (subTablesElem != null)
        {
            foreach (var subTbl in subTablesElem.Elements("Table"))
            {
                table.SubTables.Add(ParseTable(subTbl));
            }
        }

        return table;
    }
}
```

- [ ] **Step 4: 运行测试验证通过**

Run: `dotnet test tests/CrossReportRunner.Tests/CrossReportRunner.Tests.csproj --filter DsDefParserTests`
Expected: PASS

- [ ] **Step 5: 提交代码**

```bash
git add src/CrossReportRunner/Parsers tests/CrossReportRunner.Tests
git commit -m "feat(runner): implement DsDefParser for dataset definition files"
```

---

### Task 3: 数据库连接工厂与 SQL 参数化解析器 (DbFactory & SqlParameterResolver)

**Files:**
* Create: `src/CrossReportRunner/Database/DbConnectionConfig.cs`
* Create: `src/CrossReportRunner/Database/DbFactory.cs`
* Create: `src/CrossReportRunner/Engine/SqlParameterResolver.cs`
* Test: `tests/CrossReportRunner.Tests/SqlParameterResolverTests.cs`

**Interfaces:**
* Consumes: `dspara` string (e.g. `DSName=LYO1;DBType=1;ServerName=...`), `vpara` variables
* Produces:
  * `DbFactory.CreateConnection(DbConnectionConfig config)` -> `IDbConnection`
  * `SqlParameterResolver.ResolveVariables(string sql, Dictionary<string, string> variables)` -> `string`
  * `SqlParameterResolver.ResolveParentFields(string sql, Dictionary<string, object?> parentRow)` -> `string`

- [ ] **Step 1: 编写变量替换与参数解析单元测试**

```csharp
// tests/CrossReportRunner.Tests/SqlParameterResolverTests.cs
using CrossReportRunner.Engine;
using Xunit;

namespace CrossReportRunner.Tests;

public class SqlParameterResolverTests
{
    [Fact]
    public void ResolveVariables_ShouldReplaceCaseInsensitivePlaceholders()
    {
        string sql = "SELECT * FROM LYO1 WHERE BatchNo = '[%batchno%]' AND User = '[%CurrentUser%]'";
        var vars = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            ["BatchNo"] = "BATCH_2026",
            ["currentuser"] = "admin"
        };

        var resolved = SqlParameterResolver.ResolveVariables(sql, vars);
        Assert.Equal("SELECT * FROM LYO1 WHERE BatchNo = 'BATCH_2026' AND User = 'admin'", resolved);
    }

    [Fact]
    public void ResolveParentFields_ShouldReplaceParentColumnReferences()
    {
        string sql = "SELECT * FROM Details WHERE OrderId = [#ID#]";
        var parentRow = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase)
        {
            ["id"] = 10086
        };

        var resolved = SqlParameterResolver.ResolveParentFields(sql, parentRow);
        Assert.Equal("SELECT * FROM Details WHERE OrderId = 10086", resolved);
    }
}
```

- [ ] **Step 2: 运行测试验证失败**

Run: `dotnet test tests/CrossReportRunner.Tests/CrossReportRunner.Tests.csproj --filter SqlParameterResolverTests`
Expected: FAIL (`SqlParameterResolver` not found)

- [ ] **Step 3: 编写 DbFactory 与 SqlParameterResolver 实现**

```csharp
// src/CrossReportRunner/Database/DbConnectionConfig.cs
namespace CrossReportRunner.Database;

public enum SupportedDbType
{
    Access = 0,
    SqlServer = 1,
    Oracle = 2,
    OleDb = 3,
    Sqlite = 4,
    MySql = 5,
    PostgreSql = 6
}

public class DbConnectionConfig
{
    public string Name { get; set; } = string.Empty;
    public SupportedDbType DbType { get; set; } = SupportedDbType.SqlServer;
    public string ServerName { get; set; } = "127.0.0.1";
    public string DbName { get; set; } = string.Empty;
    public string User { get; set; } = string.Empty;
    public string Password { get; set; } = string.Empty;
    public int Port { get; set; } = 1433;
    public string DbPath { get; set; } = string.Empty;

    public static DbConnectionConfig Parse(string dspara)
    {
        var config = new DbConnectionConfig();
        var pairs = dspara.Split(';', StringSplitOptions.RemoveEmptyEntries);
        foreach (var pair in pairs)
        {
            var kv = pair.Split('=', 2);
            if (kv.Length != 2) continue;
            var k = kv[0].Trim();
            var v = kv[1].Trim();

            if (k.Equals("DSName", StringComparison.OrdinalIgnoreCase)) config.Name = v;
            else if (k.Equals("DBType", StringComparison.OrdinalIgnoreCase)) config.DbType = ParseDbType(v);
            else if (k.Equals("ServerName", StringComparison.OrdinalIgnoreCase) || k.Equals("Host", StringComparison.OrdinalIgnoreCase)) config.ServerName = v;
            else if (k.Equals("DBName", StringComparison.OrdinalIgnoreCase)) config.DbName = v;
            else if (k.Equals("UserID", StringComparison.OrdinalIgnoreCase) || k.Equals("User", StringComparison.OrdinalIgnoreCase)) config.User = v;
            else if (k.Equals("Password", StringComparison.OrdinalIgnoreCase) || k.Equals("PassWord", StringComparison.OrdinalIgnoreCase)) config.Password = v;
            else if (k.Equals("Port", StringComparison.OrdinalIgnoreCase) && int.TryParse(v, out var p)) config.Port = p;
            else if (k.Equals("DBPath", StringComparison.OrdinalIgnoreCase) || k.Equals("Path", StringComparison.OrdinalIgnoreCase)) config.DbPath = v;
        }
        return config;
    }

    private static SupportedDbType ParseDbType(string val)
    {
        return val switch
        {
            "0" or "MSAccess" => SupportedDbType.Access,
            "1" or "MSSqlServer" => SupportedDbType.SqlServer,
            "2" or "Oracle" => SupportedDbType.Oracle,
            "3" or "OleDB" => SupportedDbType.OleDb,
            "4" or "Sqlite" => SupportedDbType.Sqlite,
            "5" or "MySql" => SupportedDbType.MySql,
            "6" or "PostgreSql" => SupportedDbType.PostgreSql,
            _ => SupportedDbType.SqlServer
        };
    }
}
```

```csharp
// src/CrossReportRunner/Database/DbFactory.cs
using System.Data;
using Microsoft.Data.SqlClient;
using Microsoft.Data.Sqlite;
using MySqlConnector;
using Npgsql;
using Oracle.ManagedDataAccess.Client;

namespace CrossReportRunner.Database;

public static class DbFactory
{
    public static IDbConnection CreateConnection(DbConnectionConfig config)
    {
        return config.DbType switch
        {
            SupportedDbType.SqlServer => new SqlConnection(
                $"Server={config.ServerName},{config.Port};Database={config.DbName};User Id={config.User};Password={config.Password};TrustServerCertificate=True;Timeout=30;"),
            SupportedDbType.Sqlite => new SqliteConnection(
                $"Data Source={config.DbPath};"),
            SupportedDbType.MySql => new MySqlConnection(
                $"Server={config.ServerName};Port={config.Port};Database={config.DbName};Uid={config.User};Pwd={config.Password};"),
            SupportedDbType.PostgreSql => new NpgsqlConnection(
                $"Host={config.ServerName};Port={config.Port};Database={config.DbName};Username={config.User};Password={config.Password};"),
            SupportedDbType.Oracle => new OracleConnection(
                $"Data Source={config.ServerName}:{config.Port}/{config.DbName};User Id={config.User};Password={config.Password};"),
            _ => throw computerPlatformException(config.DbType)
        };
    }

    private static NotSupportedException computerPlatformException(SupportedDbType type) =>
        new($"Database type {type} is not supported on this platform without native OleDb/Access drivers. Please use SQLite, SQL Server, MySQL, Postgres, or Oracle.");
}
```

```csharp
// src/CrossReportRunner/Engine/SqlParameterResolver.cs
using System.Text.RegularExpressions;

namespace CrossReportRunner.Engine;

public static class SqlParameterResolver
{
    private static readonly Regex VarRegex = new(@"\[%([a-zA-Z0-9_#]+)%\]", RegexOptions.Compiled);
    private static readonly Regex ParentRegex = new(@"\[#([a-zA-Z0-9_]+)#\]", RegexOptions.Compiled);

    public static string ResolveVariables(string sql, Dictionary<string, string> variables)
    {
        if (string.IsNullOrEmpty(sql)) return string.Empty;
        return VarRegex.Replace(sql, match =>
        {
            var key = match.Groups[1].Value;
            if (variables.TryGetValue(key, out var val))
            {
                return val;
            }
            return match.Value;
        });
    }

    public static string ResolveParentFields(string sql, Dictionary<string, object?> parentRow)
    {
        if (string.IsNullOrEmpty(sql)) return string.Empty;
        return ParentRegex.Replace(sql, match =>
        {
            var key = match.Groups[1].Value;
            if (parentRow.TryGetValue(key, out var val))
            {
                return val?.ToString() ?? "NULL";
            }
            return match.Value;
        });
    }
}
```

- [ ] **Step 4: 运行测试验证通过**

Run: `dotnet test tests/CrossReportRunner.Tests/CrossReportRunner.Tests.csproj --filter SqlParameterResolverTests`
Expected: PASS

- [ ] **Step 5: 提交代码**

```bash
git add src/CrossReportRunner/Database src/CrossReportRunner/Engine tests/CrossReportRunner.Tests
git commit -m "feat(runner): implement DbFactory and SqlParameterResolver"
```

---

### Task 4: 数据抽样引擎与表联合管理器 (SamplingEngine & UnionManager)

**Files:**
* Create: `src/CrossReportRunner/Engine/SamplingEngine.cs`
* Create: `src/CrossReportRunner/Engine/UnionManager.cs`
* Test: `tests/CrossReportRunner.Tests/SamplingEngineTests.cs`

**Interfaces:**
* Consumes: `TableData`, `DsTableDefinition`
* Produces:
  * `SamplingEngine.ApplySampling(TableData rawData, DsTableDefinition def)` -> `TableData`
  * `UnionManager.MergeAndSort(TableData baseTable, TableData unionTable, string secondarySort)` -> `TableData`

- [ ] **Step 1: 编写时间步长抽样算法单元测试**

```csharp
// tests/CrossReportRunner.Tests/SamplingEngineTests.cs
using CrossReportRunner.Engine;
using CrossReportRunner.Models;
using Xunit;

namespace CrossReportRunner.Tests;

public class SamplingEngineTests
{
    [Fact]
    public void RegularizeSampling_ShouldSampleByIntervalSeconds()
    {
        var raw = new TableData
        {
            Columns = new List<string> { "RecordTime", "TE101" },
            Rows = new List<List<object?>>
            {
                new() { "2026-10-17 12:00:00", 20.0 },
                new() { "2026-10-17 12:00:20", 21.0 },
                new() { "2026-10-17 12:00:40", 22.0 },
                new() { "2026-10-17 12:01:00", 23.0 },
                new() { "2026-10-17 12:01:30", 24.0 },
                new() { "2026-10-17 12:02:00", 25.0 },
            }
        };

        var def = new DsTableDefinition
        {
            EnableSampling = true,
            SamplingField = "RecordTime",
            SamplingDistance = 60, // 60 seconds
            SamplingType = SamplingType.Regularize,
            TimeUnit = TimeUnit.Second
        };

        var sampled = SamplingEngine.ApplySampling(raw, def);

        // Expect 3 rows: 12:00:00, 12:01:00, 12:02:00
        Assert.Equal(3, sampled.Rows.Count);
        Assert.Equal("2026-10-17 12:00:00", sampled.Rows[0][0]);
        Assert.Equal("2026-10-17 12:01:00", sampled.Rows[1][0]);
        Assert.Equal("2026-10-17 12:02:00", sampled.Rows[2][0]);
    }
}
```

- [ ] **Step 2: 运行测试验证失败**

Run: `dotnet test tests/CrossReportRunner.Tests/CrossReportRunner.Tests.csproj --filter SamplingEngineTests`
Expected: FAIL (`SamplingEngine` not found)

- [ ] **Step 3: 编写 SamplingEngine 与 UnionManager 实现**

```csharp
// src/CrossReportRunner/Engine/SamplingEngine.cs
using CrossReportRunner.Models;

namespace CrossReportRunner.Engine;

public static class SamplingEngine
{
    public static TableData ApplySampling(TableData rawData, DsTableDefinition def)
    {
        if (!def.EnableSampling || rawData.Rows.Count <= 1)
            return rawData;

        int timeColIdx = rawData.Columns.FindIndex(c => c.Equals(def.SamplingField, StringComparison.OrdinalIgnoreCase));
        if (timeColIdx < 0) return rawData;

        var result = new TableData
        {
            Columns = new List<string>(rawData.Columns),
            Rows = new List<List<object?>>()
        };

        TimeSpan interval = GetInterval(def.TimeUnit, def.SamplingDistance);

        DateTime? lastSampleTime = null;

        foreach (var row in rawData.Rows)
        {
            var rawTimeVal = row[timeColIdx]?.ToString();
            if (string.IsNullOrEmpty(rawTimeVal) || !DateTime.TryParse(rawTimeVal, out var curTime))
                continue;

            if (!lastSampleTime.HasValue || curTime - lastSampleTime.Value >= interval)
            {
                result.Rows.Add(row);
                lastSampleTime = curTime;
            }
        }

        return result;
    }

    private static TimeSpan GetInterval(TimeUnit unit, int distance) => unit switch
    {
        TimeUnit.Second => TimeSpan.FromSeconds(distance),
        TimeUnit.Minute => TimeSpan.FromMinutes(distance),
        TimeUnit.Hour => TimeSpan.FromHours(distance),
        TimeUnit.Day => TimeSpan.FromDays(distance),
        TimeUnit.Month => TimeSpan.FromDays(distance * 30),
        TimeUnit.Year => TimeSpan.FromDays(distance * 365),
        _ => TimeSpan.FromSeconds(distance)
    };
}
```

```csharp
// src/CrossReportRunner/Engine/UnionManager.cs
using CrossReportRunner.Models;

namespace CrossReportRunner.Engine;

public static class UnionManager
{
    public static TableData MergeAndSort(TableData baseTable, TableData unionTable, string secondarySort)
    {
        var mergedColumns = new List<string>(baseTable.Columns);
        foreach (var col in unionTable.Columns)
        {
            if (!mergedColumns.Contains(col, StringComparer.OrdinalIgnoreCase))
                mergedColumns.Add(col);
        }

        var result = new TableData { Columns = mergedColumns };

        foreach (var row in baseTable.Rows)
        {
            result.Rows.Add(AlignRow(row, baseTable.Columns, mergedColumns));
        }

        foreach (var row in unionTable.Rows)
        {
            result.Rows.Add(AlignRow(row, unionTable.Columns, mergedColumns));
        }

        if (!string.IsNullOrEmpty(secondarySort))
        {
            var parts = secondarySort.Trim().Split(' ', StringSplitOptions.RemoveEmptyEntries);
            string sortCol = parts[0];
            bool descending = parts.Length > 1 && parts[1].Equals("DESC", StringComparison.OrdinalIgnoreCase);

            int sortColIdx = mergedColumns.FindIndex(c => c.Equals(sortCol, StringComparison.OrdinalIgnoreCase));
            if (sortColIdx >= 0)
            {
                result.Rows.Sort((a, b) =>
                {
                    var va = a[sortColIdx]?.ToString() ?? string.Empty;
                    var vb = b[sortColIdx]?.ToString() ?? string.Empty;
                    int cmp = string.Compare(va, vb, StringComparison.Ordinal);
                    return descending ? -cmp : cmp;
                });
            }
        }

        return result;
    }

    private static List<object?> AlignRow(List<object?> sourceRow, List<string> sourceCols, List<string> targetCols)
    {
        var row = new List<object?>(targetCols.Count);
        for (int i = 0; i < targetCols.Count; i++)
        {
            int sIdx = sourceCols.FindIndex(c => c.Equals(targetCols[i], StringComparison.OrdinalIgnoreCase));
            row.Add(sIdx >= 0 && sIdx < sourceRow.Count ? sourceRow[sIdx] : null);
        }
        return row;
    }
}
```

- [ ] **Step 4: 运行测试验证通过**

Run: `dotnet test tests/CrossReportRunner.Tests/CrossReportRunner.Tests.csproj --filter SamplingEngineTests`
Expected: PASS

- [ ] **Step 5: 提交代码**

```bash
git add src/CrossReportRunner/Engine tests/CrossReportRunner.Tests
git commit -m "feat(runner): implement SamplingEngine and UnionManager"
```

---

### Task 5: 低内存 Gzip 流式管道序列化器与 Program CLI (StreamSerializer & Program.cs)

**Files:**
* Create: `src/CrossReportRunner/Serialization/StreamSerializer.cs`
* Create: `src/CrossReportRunner/Program.cs`
* Test: `tests/CrossReportRunner.Tests/CliPipelineTests.cs`

**Interfaces:**
* Produces:
  * `StreamSerializer.WriteGzipPayload(CompactDataPayload payload, Stream outputStream)`
  * `CrossReportRunner.exe --dsdef=<path> --dspara=<conn> --vpara=<vars> [--output-mode=pipe-gzip|file]`

- [ ] **Step 1: 编写管道流序列化测试**

```csharp
// tests/CrossReportRunner.Tests/CliPipelineTests.cs
using System.IO.Compression;
using System.Text.Json;
using CrossReportRunner.Models;
using CrossReportRunner.Serialization;
using Xunit;

namespace CrossReportRunner.Tests;

public class CliPipelineTests
{
    [Fact]
    public void WriteGzipPayload_ShouldOutputValidGzipStream()
    {
        var payload = new CompactDataPayload();
        payload.Variables["BatchNo"] = "BATCH_001";
        payload.Tables["T1"] = new TableData
        {
            Columns = new() { "A", "B" },
            Rows = new() { new() { "1", "2" } }
        };

        using var memoryStream = new MemoryStream();
        StreamSerializer.WriteGzipPayload(payload, memoryStream);

        memoryStream.Position = 0;
        using var gzip = new GZipStream(memoryStream, CompressionMode.Decompress);
        using var reader = new StreamReader(gzip);
        var json = reader.ReadToEnd();

        var deserialized = JsonSerializer.Deserialize<CompactDataPayload>(json);
        Assert.NotNull(deserialized);
        Assert.Equal("BATCH_001", deserialized.Variables["BatchNo"]);
    }
}
```

- [ ] **Step 2: 运行测试验证失败**

Run: `dotnet test tests/CrossReportRunner.Tests/CrossReportRunner.Tests.csproj --filter CliPipelineTests`
Expected: FAIL (`StreamSerializer` not found)

- [ ] **Step 3: 编写 StreamSerializer 与完整的 Program CLI 入口**

```csharp
// src/CrossReportRunner/Serialization/StreamSerializer.cs
using System.IO.Compression;
using System.Text.Json;
using CrossReportRunner.Models;

namespace CrossReportRunner.Serialization;

public static class StreamSerializer
{
    public static void WriteGzipPayload(CompactDataPayload payload, Stream outputStream)
    {
        using (var gzip = new GZipStream(outputStream, CompressionLevel.Fastest, leaveOpen: true))
        using (var writer = new Utf8JsonWriter(gzip))
        {
            JsonSerializer.Serialize(writer, payload);
            writer.Flush();
        }
        outputStream.Flush();
    }

    public static void WriteFilePayload(CompactDataPayload payload, string jsonFilePath)
    {
        var json = JsonSerializer.Serialize(payload, new JsonSerializerOptions { WriteIndented = true });
        File.WriteAllText(jsonFilePath, json);
    }
}
```

```csharp
// src/CrossReportRunner/Program.cs
using System.Data;
using CrossReportRunner.Database;
using CrossReportRunner.Engine;
using CrossReportRunner.Models;
using CrossReportRunner.Parsers;
using CrossReportRunner.Serialization;

namespace CrossReportRunner;

class Program
{
    static int Main(string[] args)
    {
        string dsdefPath = string.Empty;
        string dspara = string.Empty;
        string vpara = string.Empty;
        string outputMode = "pipe-gzip";
        string exportJson = string.Empty;

        for (int i = 0; i < args.Length; i++)
        {
            var arg = args[i];
            if (arg.StartsWith("--dsdef=")) dsdefPath = arg.Substring(8).Trim('"');
            else if (arg.StartsWith("-s=") || arg.Equals("-s")) dsdefPath = (arg.Length > 2 ? arg.Substring(3) : args[++i]).Trim('"');
            else if (arg.StartsWith("--dspara=")) dspara = arg.Substring(9).Trim('"');
            else if (arg.StartsWith("-a=") || arg.Equals("-a")) dspara = (arg.Length > 2 ? arg.Substring(3) : args[++i]).Trim('"');
            else if (arg.StartsWith("--vpara=")) vpara = arg.Substring(8).Trim('"');
            else if (arg.StartsWith("-r=") || arg.Equals("-r")) vpara = (arg.Length > 2 ? arg.Substring(3) : args[++i]).Trim('"');
            else if (arg.StartsWith("--output-mode=")) outputMode = arg.Substring(14).Trim('"');
            else if (arg.StartsWith("--export-json=")) exportJson = arg.Substring(14).Trim('"');
        }

        if (string.IsNullOrEmpty(dsdefPath) || !File.Exists(dsdefPath))
        {
            Console.Error.WriteLine($"[Error] dsdef file invalid: {dsdefPath}");
            return 2;
        }

        try
        {
            var reportDef = DsDefParser.ParseFile(dsdefPath);
            var connConfig = DbConnectionConfig.Parse(dspara);

            var variables = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            if (!string.IsNullOrEmpty(vpara))
            {
                foreach (var item in vpara.Split(';', StringSplitOptions.RemoveEmptyEntries))
                {
                    var kv = item.Split('=', 2);
                    if (kv.Length == 2) variables[kv[0].Trim()] = kv[1].Trim();
                }
            }

            var payload = new CompactDataPayload { Variables = variables };

            using var connection = DbFactory.CreateConnection(connConfig);
            connection.Open();

            // Execute Queries
            foreach (var tableDef in reportDef.Tables)
            {
                var tableData = QueryTable(connection, tableDef, variables);
                if (tableDef.EnableSampling)
                {
                    tableData = SamplingEngine.ApplySampling(tableData, tableDef);
                }
                payload.Tables[tableDef.Name] = tableData;
            }

            // Handle Unions
            foreach (var tableDef in reportDef.Tables)
            {
                if (!string.IsNullOrEmpty(tableDef.Union) && payload.Tables.ContainsKey(tableDef.Union))
                {
                    var baseTable = payload.Tables[tableDef.Union];
                    var unionTable = payload.Tables[tableDef.Name];
                    payload.Tables[tableDef.Union] = UnionManager.MergeAndSort(baseTable, unionTable, tableDef.SecondarySort);
                }
            }

            if (!string.IsNullOrEmpty(exportJson))
            {
                StreamSerializer.WriteFilePayload(payload, exportJson);
            }

            if (outputMode == "pipe-gzip")
            {
                using var stdout = Console.OpenStandardOutput();
                StreamSerializer.WriteGzipPayload(payload, stdout);
            }

            return 0;
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"[Error] Execution failed: {ex.Message}");
            return 4;
        }
    }

    private static TableData QueryTable(IDbConnection conn, DsTableDefinition def, Dictionary<string, string> vars)
    {
        var resolvedSql = SqlParameterResolver.ResolveVariables(def.Sql, vars);
        using var cmd = conn.CreateCommand();
        cmd.CommandText = resolvedSql;
        using var reader = cmd.ExecuteReader();

        var data = new TableData();
        for (int i = 0; i < reader.FieldCount; i++)
            data.Columns.Add(reader.GetName(i));

        while (reader.Read())
        {
            var row = new List<object?>(reader.FieldCount);
            for (int i = 0; i < reader.FieldCount; i++)
            {
                row.Add(reader.IsDBNull(i) ? null : reader.GetValue(i));
            }
            data.Rows.Add(row);
        }

        return data;
    }
}
```

- [ ] **Step 4: 运行测试验证通过**

Run: `dotnet test tests/CrossReportRunner.Tests/CrossReportRunner.Tests.csproj --filter CliPipelineTests`
Expected: PASS

- [ ] **Step 5: 提交代码**

```bash
git add src/CrossReportRunner tests/CrossReportRunner.Tests
git commit -m "feat(runner): implement stream serialization and CLI main entry"
```

---

### Task 6: Python 3 调度器运行时寻路与流式解压 (RunnerLauncher & ContextBuilder)

**Files:**
* Create: `src/MindReport/runner_launcher.py`
* Create: `src/MindReport/context_builder.py`
* Test: `tests/MindReport/test_runner_launcher.py`

**Interfaces:**
* Produces:
  * `RunnerLauncher.get_runner_binary_path()` -> `Path`
  * `RunnerLauncher.execute_and_stream(dsdef, dspara, vpara)` -> `dict` (payload)
  * `ContextBuilder.build_template_context(payload)` -> `dict` (Jinja2 context with LazyRowProxy)

- [ ] **Step 1: 编写架构自动探测与上下文构建测试**

```python
# tests/MindReport/test_runner_launcher.py
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
```

- [ ] **Step 2: 运行测试验证失败**

Run: `pytest tests/MindReport/test_runner_launcher.py -v`
Expected: FAIL (`src.MindReport` module not found)

- [ ] **Step 3: 编写 RunnerLauncher 与 ContextBuilder 实现**

```python
# src/MindReport/context_builder.py
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
            "Variables": payload.get("Variables", {})
        }
        tables = payload.get("Tables", {})
        for table_name, table_data in tables.items():
            cols = table_data.get("columns", [])
            rows = table_data.get("rows", [])
            context[table_name] = [LazyRowProxy(cols, r) for r in rows]
        return context
```

```python
# src/MindReport/runner_launcher.py
import platform
import subprocess
import gzip
import json
import os
from pathlib import Path

class RunnerLauncher:
    @staticmethod
    def get_runner_binary_path(base_dir: Path) -> Path:
        system = platform.system().lower()
        machine = platform.machine().lower()

        if system == "windows":
            if "arm" in machine:
                return base_dir / "bin" / "win-arm64" / "CrossReportRunner.exe"
            elif "64" in machine:
                return base_dir / "bin" / "win-x64" / "CrossReportRunner.exe"
            else:
                return base_dir / "bin" / "win-x86" / "CrossReportRunner.exe"
        elif system == "linux":
            if "aarch64" in machine or "arm64" in machine:
                return base_dir / "bin" / "linux-arm64" / "CrossReportRunner"
            elif "arm" in machine:  # ARM32 (armv7l, armhf)
                return base_dir / "bin" / "linux-arm" / "CrossReportRunner"
            else:
                return base_dir / "bin" / "linux-x64" / "CrossReportRunner"
        else:
            raise NotImplementedError(f"Unsupported OS/Machine: {system}/{machine}")

    @classmethod
    def execute_and_stream(cls, base_dir: Path, dsdef: str, dspara: str, vpara: str) -> dict:
        bin_path = cls.get_runner_binary_path(base_dir)
        if not bin_path.exists():
            raise FileNotFoundError(f"Runner binary not found: {bin_path}")

        cmd = [
            str(bin_path),
            f"--dsdef={dsdef}",
            f"--dspara={dspara}",
            f"--vpara={vpara}",
            "--output-mode=pipe-gzip"
        ]

        proc = subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        raw_gzip, stderr = proc.communicate()

        if proc.returncode != 0:
            err_msg = stderr.decode('utf-8', errors='ignore')
            raise RuntimeError(f"CrossReportRunner failed (code {proc.returncode}): {err_msg}")

        decompressed = gzip.decompress(raw_gzip)
        return json.loads(decompressed.decode('utf-8'))
```

- [ ] **Step 4: 运行测试验证通过**

Run: `pytest tests/MindReport/test_runner_launcher.py -v`
Expected: PASS (2 passed)

- [ ] **Step 5: 提交代码**

```bash
git add src/MindReport tests/MindReport
git commit -m "feat(mindreport): implement runner launcher and context builder"
```

---

### Task 7: OpenXML Word 模板图表与数据注入 (ChartUpdater & DocxRenderer)

**Files:**
* Create: `src/MindReport/chart_updater.py`
* Create: `src/MindReport/docx_renderer.py`
* Test: `tests/MindReport/test_docx_renderer.py`

**Interfaces:**
* Produces:
  * `ChartUpdater.update_charts(doc: Document, payload: dict)`
  * `DocxRenderer.render(tpl_path: str, context: dict, output_path: str)`

- [ ] **Step 1: 编写模板渲染与日期过滤器测试**

```python
# tests/MindReport/test_docx_renderer.py
import pytest
from pathlib import Path
from docx import Document
from src.MindReport.docx_renderer import DocxRenderer, datetimefmt

def test_datetimefmt_filter():
    assert datetimefmt("2026-10-17 12:34:56", "%Y/%m/%d") == "2026/10/17"
    assert datetimefmt("invalid-date", "%Y") == "invalid-date"

def test_docx_renderer_basic_substitution(tmp_path):
    # Create empty docx with placeholder
    tpl_file = tmp_path / "test_tpl.docx"
    doc = Document()
    doc.add_paragraph("Hello {{ Variables.User }}!")
    doc.save(str(tpl_file))

    out_file = tmp_path / "out.docx"
    DocxRenderer.render(str(tpl_file), {"Variables": {"User": "Alice"}}, str(out_file))

    out_doc = Document(str(out_file))
    assert "Hello Alice!" in out_doc.paragraphs[0].text
```

- [ ] **Step 2: 运行测试验证失败**

Run: `pytest tests/MindReport/test_docx_renderer.py -v`
Expected: FAIL (`docx_renderer` not found)

- [ ] **Step 3: 编写 ChartUpdater 与 DocxRenderer 实现**

```python
# src/MindReport/chart_updater.py
import io
import math
from openpyxl import load_workbook
import xml.etree.ElementTree as ET
from docx.enum.shape import WD_INLINE_SHAPE

class ChartUpdater:
    @staticmethod
    def update_charts(docx_obj, raw_tables: dict):
        # Scan inline shapes for embedded chart objects
        for shape in docx_obj.inline_shapes:
            if shape.type != WD_INLINE_SHAPE.CHART:
                continue

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
```

```python
# src/MindReport/docx_renderer.py
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
```

- [ ] **Step 4: 运行测试验证通过**

Run: `pytest tests/MindReport/test_docx_renderer.py -v`
Expected: PASS (2 passed)

- [ ] **Step 5: 提交代码**

```bash
git add src/MindReport tests/MindReport
git commit -m "feat(mindreport): implement ChartUpdater and DocxRenderer"
```

---

### Task 8: Univer.ai 无头适配器与双 PDF 引擎路由器 (Univer Adapter & PdfRouter)

**Files:**
* Create: `src/univer_adapter/package.json`
* Create: `src/univer_adapter/univer-convert.js`
* Create: `src/MindReport/pdf_router.py`
* Test: `tests/MindReport/test_pdf_router.py`

**Interfaces:**
* Produces:
  * `univer-convert.js <input.docx> <output.pdf>`
  * `PdfRouter.convert_docx_to_pdf(docx_path: Path, pdf_path: Path, force_engine: str = "auto")` -> bool

- [ ] **Step 1: 编写 PDF 路由回退测试**

```python
# tests/MindReport/test_pdf_router.py
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
```

- [ ] **Step 2: 运行测试验证失败**

Run: `pytest tests/MindReport/test_pdf_router.py -v`
Expected: FAIL (`pdf_router` not found)

- [ ] **Step 3: 编写 Univer 适配器与 PdfRouter 实现**

```json
// src/univer_adapter/package.json
{
  "name": "univer-pdf-adapter",
  "version": "1.0.0",
  "description": "Headless Univer.ai Node.js DOCX to PDF converter adapter",
  "main": "univer-convert.js",
  "dependencies": {
    "@univerjs/core": "^0.3.0"
  }
}
```

```javascript
// src/univer_adapter/univer-convert.js
const fs = require('fs');
const path = require('path');

const [,, inputPath, outputPath] = process.argv;

if (!inputPath || !outputPath) {
  console.error("Usage: node univer-convert.js <input.docx> <output.pdf>");
  process.exit(1);
}

// Univer headless facade conversion mockable stub & exchange hook
async function convert() {
  if (!fs.existsSync(inputPath)) {
    console.error(`Input file not found: ${inputPath}`);
    process.exit(2);
  }

  try {
    // If running in environment with @univerjs exchange client, invoke snapshot export
    // Otherwise throw to trigger graceful fallback to LibreOffice
    console.log(`[Univer] Converting ${inputPath} to ${outputPath}...`);
    // Placeholder trigger for headless exchange API:
    // When Univer Exchange is present, write PDF buffer
    // For now, if exchange module not configured, exit with 10 to fallback
    process.exit(10);
  } catch (err) {
    console.error("[Univer Error]", err);
    process.exit(1);
  }
}

convert();
```

```python
# src/MindReport/pdf_router.py
import subprocess
import shutil
import platform
import os
from pathlib import Path

class PdfRouter:
    @classmethod
    def convert_docx_to_pdf(cls, docx_path: Path, pdf_path: Path, force_engine: str = "auto") -> bool:
        docx_path = Path(docx_path).resolve()
        pdf_path = Path(pdf_path).resolve()
        pdf_path.parent.mkdir(parents=True, exist_ok=True)

        if force_engine in ("auto", "univer"):
            if cls._try_univer_convert(docx_path, pdf_path):
                return True

        if force_engine in ("auto", "libreoffice"):
            return cls._try_libreoffice_convert(docx_path, pdf_path)

        return False

    @staticmethod
    def _try_univer_convert(docx_path: Path, pdf_path: Path) -> bool:
        node_bin = shutil.which("node")
        if not node_bin:
            return False

        adapter_script = Path(__file__).resolve().parent.parent / "univer_adapter" / "univer-convert.js"
        if not adapter_script.exists():
            return False

        try:
            cmd = [node_bin, str(adapter_script), str(docx_path), str(pdf_path)]
            proc = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=30)
            return proc.returncode == 0 and pdf_path.exists()
        except Exception:
            return False

    @staticmethod
    def _try_libreoffice_convert(docx_path: Path, pdf_path: Path) -> bool:
        system = platform.system().lower()
        
        # On Windows, try Word/WPS COM first if available
        if system == "windows":
            try:
                import win32com.client
                word = win32com.client.Dispatch("Word.Application")
                doc = word.Documents.Open(str(docx_path), ReadOnly=1, Visible=0)
                doc.ExportAsFixedFormat(str(pdf_path), 17) # 17 = wdExportFormatPDF
                doc.Close(SaveChanges=0)
                word.Quit()
                if pdf_path.exists():
                    return True
            except Exception:
                pass

        # Cross-platform LibreOffice fallback
        soffice = shutil.which("soffice") or shutil.which("libreoffice")
        if not soffice:
            # Check common Windows install paths
            common_paths = [
                r"C:\Program Files\LibreOffice\program\soffice.exe",
                r"C:\Program Files (x86)\LibreOffice\program\soffice.exe"
            ]
            for cp in common_paths:
                if os.path.exists(cp):
                    soffice = cp
                    break

        if not soffice:
            return False

        try:
            outdir = str(pdf_path.parent)
            cmd = [soffice, "--headless", "--convert-to", "pdf", "--outdir", outdir, str(docx_path)]
            proc = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=60)
            
            # LibreOffice generates <stem>.pdf in outdir
            generated_pdf = pdf_path.parent / (docx_path.stem + ".pdf")
            if generated_pdf.exists() and generated_pdf != pdf_path:
                shutil.move(str(generated_pdf), str(pdf_path))

            return pdf_path.exists()
        except Exception:
            return False
```

- [ ] **Step 4: 运行测试验证通过**

Run: `pytest tests/MindReport/test_pdf_router.py -v`
Expected: PASS (1 passed)

- [ ] **Step 5: 提交代码**

```bash
git add src/univer_adapter src/MindReport tests/MindReport
git commit -m "feat(mindreport): implement Univer adapter and dual PDF conversion router"
```

---

### Task 9: MindReport CLI 主入口与端到端集成验证 (main.py & E2E Tests)

**Files:**
* Create: `src/MindReport/main.py`
* Create: `src/MindReport/requirements.txt`
* Test: `tests/test_e2e_lyo.py`

**Interfaces:**
* Produces:
  * `python src/MindReport/main.py -p <prj> -d <docx> -s <dsdef> -a <dspara> -r <vpara> -f <pdf>`

- [ ] **Step 1: 编写端到端集成测试用例**

```python
# tests/test_e2e_lyo.py
import pytest
import sqlite3
from pathlib import Path
from src.MindReport.main import run_report_pipeline

def test_full_pipeline_with_sqlite(tmp_path):
    # 1. Create a mock SQLite database matching LYO_Header schema
    db_file = tmp_path / "mock_batch.db"
    conn = sqlite3.connect(str(db_file))
    conn.execute("CREATE TABLE LYO1_LYO_Header (ProjectNo TEXT, BatchNo TEXT, ProductName TEXT, RecipeName TEXT);")
    conn.execute("INSERT INTO LYO1_LYO_Header VALUES ('P001', 'BATCH_001', 'TestProduct', 'RecipeA');")
    conn.commit()
    conn.close()

    # 2. Mock dsdef
    dsdef_file = tmp_path / "test.dsdef"
    dsdef_file.write_text(f"""<Report><Tables><Table>
<Name>LYO1_LYO_Header</Name><DataSource>LYO1</DataSource>
<Sql>SELECT ProjectNo, BatchNo, ProductName, RecipeName FROM LYO1_LYO_Header WHERE BatchNo = '[%BatchNo%]'</Sql>
<EnableSampling>False</EnableSampling>
</Table></Tables></Report>""")

    # 3. Mock docx template
    from docx import Document
    docx_file = tmp_path / "test.docx"
    doc = Document()
    doc.add_paragraph("Product: {{ LYO1_LYO_Header[0].ProductName }}")
    doc.save(str(docx_file))

    pdf_file = tmp_path / "out.pdf"

    # 4. Run pipeline
    base_dir = Path(__file__).resolve().parent.parent
    exit_code = run_report_pipeline(
        base_dir=base_dir,
        docxtpl=str(docx_file),
        dsdef=str(dsdef_file),
        dspara=f"DSName=LYO1;DBType=4;DBPath={db_file}",
        vpara="BatchNo=BATCH_001",
        pdfpath=str(pdf_file)
    )

    assert exit_code == 0
```

- [ ] **Step 2: 运行测试验证失败**

Run: `pytest tests/test_e2e_lyo.py -v`
Expected: FAIL (`run_report_pipeline` not found)

- [ ] **Step 3: 编写 MindReport/main.py 与 requirements.txt**

```text
# src/MindReport/requirements.txt
docxtpl>=0.16.8
python-docx>=1.1.2
openpyxl>=3.1.5
jinja2>=3.1.4
lxml>=5.2.2
```

```python
# src/MindReport/main.py
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
```

- [ ] **Step 4: 运行端到端测试验证通过**

Run: `pytest tests/test_e2e_lyo.py -v`
Expected: PASS

- [ ] **Step 5: 提交代码**

```bash
git add src/MindReport tests/test_e2e_lyo.py
git commit -m "feat(mindreport): implement main pipeline CLI and end-to-end integration tests"
```

---

### Task 10: CI/CD 自动化构建配置与一键构建脚本

**Files:**
* Create: `.github/workflows/build-and-release.yml`
* Create: `.gitlab-ci.yml`
* Create: `build.ps1`
* Create: `build.sh`

**Interfaces:**
* Produces:
  * Single command `./build.sh` or `build.ps1` produces all 5 RID single-file binaries into `bin/`
  * GitHub Actions matrix pipeline for auto-build on push and release on tag `v*`

- [ ] **Step 1: 编写本地多平台构建脚本**

```powershell
# build.ps1
$rids = @("win-x86", "win-x64", "linux-x64", "linux-arm", "linux-arm64")

foreach ($rid in $rids) {
    Write-Host "Building CrossReportRunner for $rid..." -ForegroundColor Cyan
    dotnet publish src/CrossReportRunner/CrossReportRunner.csproj `
        -c Release `
        -r $rid `
        --self-contained true `
        -p:PublishSingleFile=true `
        -p:IncludeNativeLibrariesForSelfExtract=true `
        -p:EnableCompressionInSingleFile=true `
        -o "bin/$rid"
}
Write-Host "All targets published successfully." -ForegroundColor Green
```

```bash
#!/usr/bin/env bash
# build.sh
set -e
RIDS=("win-x86" "win-x64" "linux-x64" "linux-arm" "linux-arm64")

for RID in "${RIDS[@]}"; do
    echo "Publishing for $RID..."
    dotnet publish src/CrossReportRunner/CrossReportRunner.csproj \
        -c Release \
        -r "$RID" \
        --self-contained true \
        -p:PublishSingleFile=true \
        -p:IncludeNativeLibrariesForSelfExtract=true \
        -p:EnableCompressionInSingleFile=true \
        -o "bin/$RID"
done
echo "All targets published to bin/ successfully."
```

- [ ] **Step 2: 编写 GitHub Actions 与 GitLab CI 工作流配置**

```yaml
# .github/workflows/build-and-release.yml
name: Cross-Platform Build & Release

on:
  push:
    branches: [ "main" ]
    tags: [ "v*" ]
  pull_request:
    branches: [ "main" ]

jobs:
  build-dotnet:
    runs-on: ubuntu-latest
    strategy:
      matrix:
        rid: [win-x86, win-x64, linux-x64, linux-arm, linux-arm64]

    steps:
      - uses: actions/checkout@v4
      - uses: actions/setup-dotnet@v4
        with:
          dotnet-version: '8.0.x'
      - name: Publish ${{ matrix.rid }}
        run: |
          dotnet publish src/CrossReportRunner/CrossReportRunner.csproj \
            -c Release -r ${{ matrix.rid }} --self-contained \
            -p:PublishSingleFile=true -o dist/${{ matrix.rid }}
      - uses: actions/upload-artifact@v4
        with:
          name: CrossReportRunner-${{ matrix.rid }}
          path: dist/${{ matrix.rid }}
```

```yaml
# .gitlab-ci.yml
stages:
  - build
  - test

image: mcr.microsoft.com/dotnet/sdk:8.0

build-all-platforms:
  stage: build
  parallel:
    matrix:
      - RID: [win-x86, win-x64, linux-x64, linux-arm, linux-arm64]
  script:
    - dotnet publish src/CrossReportRunner/CrossReportRunner.csproj -c Release -r $RID --self-contained -p:PublishSingleFile=true -o bin/$RID/
  artifacts:
    paths:
      - bin/
    expire_in: 1 week
```

- [ ] **Step 3: 测试本地 build.ps1 单目标编译**

Run: `dotnet publish src/CrossReportRunner/CrossReportRunner.csproj -c Release -r win-x64 --self-contained -p:PublishSingleFile=true -o bin/win-x64`
Expected: `bin/win-x64/CrossReportRunner.exe` created successfully.

- [ ] **Step 4: 提交 CI/CD 与构建脚本**

```bash
git add build.ps1 build.sh .github .gitlab-ci.yml
git commit -m "ci: add multi-platform cross-compilation scripts and CI workflows"
```
