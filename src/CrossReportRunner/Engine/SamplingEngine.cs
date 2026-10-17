using CrossReportRunner.Models;

namespace CrossReportRunner.Engine;

public static class SamplingEngine
{
    public static TableData ApplySampling(TableData rawData, DsTableDefinition def)
    {
        if (!def.EnableSampling || rawData.Rows.Count <= 1)
            return rawData;

        int timeColIdx = rawData.Columns.FindIndex(c => c.Equals(def.SamplingField, StringComparison.OrdinalIgnoreCase));
        if (timeColIdx < 0) return rawData;

        var result = new TableData
        {
            Columns = new List<string>(rawData.Columns),
            Rows = new List<List<object?>>()
        };

        TimeSpan interval = GetInterval(def.TimeUnit, def.SamplingDistance);

        DateTime? lastSampleTime = null;

        foreach (var row in rawData.Rows)
        {
            var rawTimeVal = row[timeColIdx]?.ToString();
            if (string.IsNullOrEmpty(rawTimeVal) || !DateTime.TryParse(rawTimeVal, out var curTime))
                continue;

            if (!lastSampleTime.HasValue || curTime - lastSampleTime.Value >= interval)
            {
                result.Rows.Add(row);
                lastSampleTime = curTime;
            }
        }

        return result;
    }

    private static TimeSpan GetInterval(TimeUnit unit, int distance) => unit switch
    {
        TimeUnit.Second => TimeSpan.FromSeconds(distance),
        TimeUnit.Minute => TimeSpan.FromMinutes(distance),
        TimeUnit.Hour => TimeSpan.FromHours(distance),
        TimeUnit.Day => TimeSpan.FromDays(distance),
        TimeUnit.Month => TimeSpan.FromDays(distance * 30),
        TimeUnit.Year => TimeSpan.FromDays(distance * 365),
        _ => TimeSpan.FromSeconds(distance)
    };
}
