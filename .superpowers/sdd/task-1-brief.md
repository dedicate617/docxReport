# Task 1 Brief: .NET 8 工程脚手架与数据模型

**Files:**
* Create: `src/CrossReportRunner/CrossReportRunner.csproj`
* Create: `src/CrossReportRunner/Models/DsDefModels.cs`
* Create: `src/CrossReportRunner/Models/CompactDataPayload.cs`
* Test: `tests/CrossReportRunner.Tests/CrossReportRunner.Tests.csproj`
* Test: `tests/CrossReportRunner.Tests/ModelSerializationTests.cs`

**Global Constraints:**
* Target Framework: .NET 8 (`net8.0`)
* Single-file self-contained ready: `PublishSingleFile=true`, `SelfContained=true`, `IncludeNativeLibrariesForSelfExtract=true`, `EnableCompressionInSingleFile=true`
* Database dependencies: `Microsoft.Data.SqlClient` (v5.2.2), `Microsoft.Data.Sqlite` (v8.0.10), `MySqlConnector` (v2.3.7), `Npgsql` (v8.0.5), `Oracle.ManagedDataAccess.Core` (v23.6.0)

**Interfaces:**
* Produces:
  * `DsReportDefinition`: `.dsdef` root object model
  * `DsTableDefinition`: table query definition (Sql, Sampling, Union, Fields, RelateFields, SecondarySort)
  * `CompactDataPayload`: compact 2D payload containing `Dictionary<string, TableData> Tables` and `Dictionary<string, string> Variables`
  * `TableData`: `{ List<string> Columns, List<List<object?>> Rows }`

**Steps to execute:**
1. Create `tests/CrossReportRunner.Tests/CrossReportRunner.Tests.csproj` and `tests/CrossReportRunner.Tests/ModelSerializationTests.cs`.
2. Run test to verify it fails (`dotnet test tests/CrossReportRunner.Tests/CrossReportRunner.Tests.csproj`).
3. Create `src/CrossReportRunner/CrossReportRunner.csproj`, `src/CrossReportRunner/Models/CompactDataPayload.cs`, and `src/CrossReportRunner/Models/DsDefModels.cs`.
4. Run test to verify it passes (`dotnet test tests/CrossReportRunner.Tests/CrossReportRunner.Tests.csproj`).
5. Commit work with `git commit -m "feat(runner): initialize .NET 8 project and data models"`.
