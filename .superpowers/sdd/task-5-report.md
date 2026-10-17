# Task 5 Completion Report: 低内存 Gzip 流式管道序列化器与 Program CLI

## Status
DONE

## Summary of Changes
1. **流式管道序列化器 (StreamSerializer)**:
   - `src/CrossReportRunner/Serialization/StreamSerializer.cs`:
     - `WriteGzipPayload`: 包装 `GZipStream` 与 `Utf8JsonWriter`，将 `CompactDataPayload` 直接流式写出到标准输出或任意 `Stream`，避免中间大字符串/全量缓冲。
     - `WriteFilePayload`: 支持导出格式化 JSON 到指定文件。

2. **跨平台可执行主入口 (Program CLI)**:
   - `src/CrossReportRunner/Program.cs`:
     - 命令行参数解析：`--dsdef` (或 `-s`), `--dspara` (或 `-a`), `--vpara` (或 `-r`), `--output-mode`, `--export-json`。
     - 集成：DsDefParser 解析 -> DbFactory 建立数据库连接 -> 执行 SQL 并解析变量 -> 应用 SamplingEngine 抽样 -> 应用 UnionManager 联合与排序 -> StreamSerializer 流式 Gzip 输出。

3. **单元测试与验证**:
   - `tests/CrossReportRunner.Tests/CliPipelineTests.cs`:
     - 验证 `WriteGzipPayload_ShouldOutputValidGzipStream`：写出 Gzip 压缩二进制流，并通过 GZipStream 解压反序列化还原 Payload。
   - TDD 验证过程:
     - 初始运行: 编译失败 (RED, Serialization 未定义)。
     - 实现后运行: 单元测试全部通过 (GREEN)。
     - 全套测试: 7/7 通过 (零回归)。
