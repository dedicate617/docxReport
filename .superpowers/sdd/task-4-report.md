# Task 4 Completion Report: 数据抽样引擎与表联合管理器

## Status
DONE

## Summary of Changes
1. **抽样引擎 (SamplingEngine)**:
   - `src/CrossReportRunner/Engine/SamplingEngine.cs`:
     - 实现 `Regularize` 模式，按照设定的时间步长（秒、分、时、天、月、年）抽样。
     - 从海量工艺时序数据中提取均匀步长的数据点，极大降低数据量与网络传输/内存压力。

2. **表联合与次要排序管理器 (UnionManager)**:
   - `src/CrossReportRunner/Engine/UnionManager.cs`:
     - 实现多数据源/多表的列对齐合并（`AlignRow`）。
     - 支持 `SecondarySort` 语法（如 `BatchNo ASC`, `RecordTime DESC`），完成合并后的有序重排。

3. **单元测试与验证**:
   - `tests/CrossReportRunner.Tests/SamplingEngineTests.cs`:
     - 验证 `RegularizeSampling_ShouldSampleByIntervalSeconds`: 60 秒步长从密集秒级数据中按间隔抽样。
     - 验证 `UnionManager_ShouldMergeAndSortTables`: 对不同列集的两个表合并并按指定列升序排序。
   - TDD 验证过程:
     - 初始运行: 编译失败 (RED, SamplingEngine 未定义)。
     - 实现后运行: 单元测试全部通过 (GREEN)。
     - 全套测试: 6/6 通过 (零回归)。
