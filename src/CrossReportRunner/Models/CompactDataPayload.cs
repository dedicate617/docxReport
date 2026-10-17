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
