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
