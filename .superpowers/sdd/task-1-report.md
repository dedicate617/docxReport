# Task 1 Completion Report: .NET 8 工程脚手架与数据模型

## Status
DONE

## Summary of Changes
1. **工程与依赖配置**:
   - `src/CrossReportRunner/CrossReportRunner.csproj`: 目标框架 net8.0，引入 `Microsoft.Data.SqlClient`, `Microsoft.Data.Sqlite`, `MySqlConnector`, `Npgsql`, `Oracle.ManagedDataAccess.Core`。
   - `tests/CrossReportRunner.Tests/CrossReportRunner.Tests.csproj`: 目标框架 net8.0，引入 xUnit 和测试框架。
   - `nuget.config`: 配置 `globalPackagesFolder` 至 `Z:\E\docxReport\.nuget\packages`，解决宿主机 C 盘空间受限问题。
   - `.gitignore`: 忽略 `.tmp/`, `.nuget/`, `.dotnet/`, `bin/`, `obj/` 等。

2. **核心模型定义**:
   - `src/CrossReportRunner/Models/CompactDataPayload.cs`: 定义内存紧凑型 2D 数据负载结构 (`CompactDataPayload`, `TableData`)。
   - `src/CrossReportRunner/Models/DsDefModels.cs`: 映射 `.dsdef` 的表定义、抽样配置（Regularize/Drift）、时间单位、关系字段等对象。
   - `src/CrossReportRunner/Program.cs`: 提供初始可执行入口点。

3. **测试验证**:
   - `tests/CrossReportRunner.Tests/ModelSerializationTests.cs`: 验证 `CompactDataPayload` 的字典与 2D 二维矩阵列行序列化与反序列化逻辑。
   - 测试结果: 1/1 通过。
