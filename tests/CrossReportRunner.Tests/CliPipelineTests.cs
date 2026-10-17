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
