using System.Text.RegularExpressions;

namespace CrossReportRunner.Engine;

public static class SqlParameterResolver
{
    private static readonly Regex VarRegex = new(@"\[%([a-zA-Z0-9_#]+)%\]", RegexOptions.Compiled);
    private static readonly Regex ParentRegex = new(@"\[#([a-zA-Z0-9_]+)#\]", RegexOptions.Compiled);

    public static string ResolveVariables(string sql, Dictionary<string, string> variables)
    {
        if (string.IsNullOrEmpty(sql)) return string.Empty;
        return VarRegex.Replace(sql, match =>
        {
            var key = match.Groups[1].Value;
            if (variables.TryGetValue(key, out var val))
            {
                return val;
            }
            return match.Value;
        });
    }

    public static string ResolveParentFields(string sql, Dictionary<string, object?> parentRow)
    {
        if (string.IsNullOrEmpty(sql)) return string.Empty;
        return ParentRegex.Replace(sql, match =>
        {
            var key = match.Groups[1].Value;
            if (parentRow.TryGetValue(key, out var val))
            {
                return val?.ToString() ?? "NULL";
            }
            return match.Value;
        });
    }
}
