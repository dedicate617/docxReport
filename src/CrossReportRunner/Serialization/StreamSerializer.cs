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
        var dir = Path.GetDirectoryName(jsonFilePath);
        if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
        {
            Directory.CreateDirectory(dir);
        }
        var json = JsonSerializer.Serialize(payload, new JsonSerializerOptions { WriteIndented = true });
        File.WriteAllText(jsonFilePath, json);
    }
}
