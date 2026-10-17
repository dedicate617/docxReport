# Task 3 Completion Report: 数据库连接工厂与 SQL 参数化解析器

## Status
DONE

## Summary of Changes
1. **数据库连接工厂与配置解析**:
   - `src/CrossReportRunner/Database/DbConnectionConfig.cs`: 支持解析 `dspara` 连接字符串（`DSName`, `DBType`, `ServerName/Host`, `DBName`, `UserID/User`, `Password`, `Port`, `DBPath/Path`），映射 7 种支持的数据库类型。
   - `src/CrossReportRunner/Database/DbFactory.cs`: 使用 100% 托管 ADO.NET 客户端构建连接，支持跨平台运行：
     - SQL Server (`Microsoft.Data.SqlClient`)
     - SQLite (`Microsoft.Data.Sqlite`)
     - MySQL (`MySqlConnector`)
     - PostgreSQL (`Npgsql`)
     - Oracle (`Oracle.ManagedDataAccess.Core`)

2. **SQL 参数解析器**:
   - `src/CrossReportRunner/Engine/SqlParameterResolver.cs`:
     - `ResolveVariables`: 正则匹配 `[%VarName%]` 并进行大小写不敏感替换。
     - `ResolveParentFields`: 正则匹配 `[#ParentField#]` 并从父行关联字典中抽取并替换。

3. **单元测试与验证**:
   - `tests/CrossReportRunner.Tests/SqlParameterResolverTests.cs`:
     - 验证 `ResolveVariables` 替换大小写不敏感占位符。
     - 验证 `ResolveParentFields` 替换父行字段关联。
   - TDD 验证过程:
     - 初始运行: 编译失败 (RED, Engine 命名空间未定义)。
     - 实现后运行: 单元测试 2/2 通过 (GREEN)。
     - 全套测试: 4/4 通过 (零回归)。
