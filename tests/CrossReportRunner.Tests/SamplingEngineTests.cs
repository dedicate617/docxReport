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

    [Fact]
    public void UnionManager_ShouldMergeAndSortTables()
    {
        var tableA = new TableData
        {
            Columns = new List<string> { "BatchNo", "ValA" },
            Rows = new List<List<object?>>
            {
                new() { "B02", 20 },
                new() { "B01", 10 }
            }
        };

        var tableB = new TableData
        {
            Columns = new List<string> { "BatchNo", "ValB" },
            Rows = new List<List<object?>>
            {
                new() { "B03", 30 }
            }
        };

        var merged = UnionManager.MergeAndSort(tableA, tableB, "BatchNo ASC");
        Assert.Equal(3, merged.Rows.Count);
        Assert.Contains("ValB", merged.Columns);
        Assert.Equal("B01", merged.Rows[0][0]);
        Assert.Equal("B02", merged.Rows[1][0]);
        Assert.Equal("B03", merged.Rows[2][0]);
    }
}

