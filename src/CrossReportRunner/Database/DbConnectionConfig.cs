namespace CrossReportRunner.Database;

public enum SupportedDbType
{
    Access = 0,
    SqlServer = 1,
    Oracle = 2,
    OleDb = 3,
    Sqlite = 4,
    MySql = 5,
    PostgreSql = 6
}

public class DbConnectionConfig
{
    public string Name { get; set; } = string.Empty;
    public SupportedDbType DbType { get; set; } = SupportedDbType.SqlServer;
    public string ServerName { get; set; } = "127.0.0.1";
    public string DbName { get; set; } = string.Empty;
    public string User { get; set; } = string.Empty;
    public string Password { get; set; } = string.Empty;
    public int Port { get; set; } = 1433;
    public string DbPath { get; set; } = string.Empty;

    public static DbConnectionConfig Parse(string dspara)
    {
        var config = new DbConnectionConfig();
        var pairs = dspara.Split(';', StringSplitOptions.RemoveEmptyEntries);
        foreach (var pair in pairs)
        {
            var kv = pair.Split('=', 2);
            if (kv.Length != 2) continue;
            var k = kv[0].Trim();
            var v = kv[1].Trim();

            if (k.Equals("DSName", StringComparison.OrdinalIgnoreCase)) config.Name = v;
            else if (k.Equals("DBType", StringComparison.OrdinalIgnoreCase)) config.DbType = ParseDbType(v);
            else if (k.Equals("ServerName", StringComparison.OrdinalIgnoreCase) || k.Equals("Host", StringComparison.OrdinalIgnoreCase)) config.ServerName = v;
            else if (k.Equals("DBName", StringComparison.OrdinalIgnoreCase)) config.DbName = v;
            else if (k.Equals("UserID", StringComparison.OrdinalIgnoreCase) || k.Equals("User", StringComparison.OrdinalIgnoreCase)) config.User = v;
            else if (k.Equals("Password", StringComparison.OrdinalIgnoreCase) || k.Equals("PassWord", StringComparison.OrdinalIgnoreCase)) config.Password = v;
            else if (k.Equals("Port", StringComparison.OrdinalIgnoreCase) && int.TryParse(v, out var p)) config.Port = p;
            else if (k.Equals("DBPath", StringComparison.OrdinalIgnoreCase) || k.Equals("Path", StringComparison.OrdinalIgnoreCase)) config.DbPath = v;
        }
        return config;
    }

    private static SupportedDbType ParseDbType(string val)
    {
        return val switch
        {
            "0" or "MSAccess" => SupportedDbType.Access,
            "1" or "MSSqlServer" => SupportedDbType.SqlServer,
            "2" or "Oracle" => SupportedDbType.Oracle,
            "3" or "OleDB" => SupportedDbType.OleDb,
            "4" or "Sqlite" => SupportedDbType.Sqlite,
            "5" or "MySql" => SupportedDbType.MySql,
            "6" or "PostgreSql" => SupportedDbType.PostgreSql,
            _ => SupportedDbType.SqlServer
        };
    }
}
