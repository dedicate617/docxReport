namespace CrossReportRunner.Models;

public enum SamplingType
{
    Regularize = 0,
    Drift = 1
}

public enum TimeUnit
{
    Year = 0,
    Month = 1,
    Day = 2,
    Hour = 3,
    Minute = 4,
    Second = 5
}

public class DsFieldDefinition
{
    public string Name { get; set; } = string.Empty;
    public string FieldType { get; set; } = "0";
}

public class DsTableDefinition
{
    public string Name { get; set; } = string.Empty;
    public string DataSource { get; set; } = string.Empty;
    public string Sql { get; set; } = string.Empty;
    public string Union { get; set; } = string.Empty;
    public string RelateFields { get; set; } = string.Empty;
    public string SecondarySort { get; set; } = string.Empty;
    public bool EnableSampling { get; set; }
    public string SamplingField { get; set; } = string.Empty;
    public int SamplingDistance { get; set; } = 60;
    public SamplingType SamplingType { get; set; } = SamplingType.Regularize;
    public TimeUnit TimeUnit { get; set; } = TimeUnit.Second;
    public List<DsFieldDefinition> Fields { get; set; } = new();
    public List<DsTableDefinition> SubTables { get; set; } = new();
}

public class DsReportDefinition
{
    public List<DsTableDefinition> Tables { get; set; } = new();
}
