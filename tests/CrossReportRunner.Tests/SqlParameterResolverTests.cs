using CrossReportRunner.Engine;
using Xunit;

namespace CrossReportRunner.Tests;

public class SqlParameterResolverTests
{
    [Fact]
    public void ResolveVariables_ShouldReplaceCaseInsensitivePlaceholders()
    {
        string sql = "SELECT * FROM LYO1 WHERE BatchNo = '[%batchno%]' AND User = '[%CurrentUser%]'";
        var vars = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            ["BatchNo"] = "BATCH_2026",
            ["currentuser"] = "admin"
        };

        var resolved = SqlParameterResolver.ResolveVariables(sql, vars);
        Assert.Equal("SELECT * FROM LYO1 WHERE BatchNo = 'BATCH_2026' AND User = 'admin'", resolved);
    }

    [Fact]
    public void ResolveParentFields_ShouldReplaceParentColumnReferences()
    {
        string sql = "SELECT * FROM Details WHERE OrderId = [#ID#]";
        var parentRow = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase)
        {
            ["id"] = 10086
        };

        var resolved = SqlParameterResolver.ResolveParentFields(sql, parentRow);
        Assert.Equal("SELECT * FROM Details WHERE OrderId = 10086", resolved);
    }
}
