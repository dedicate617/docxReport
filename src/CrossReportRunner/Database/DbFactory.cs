using System.Data;
using Microsoft.Data.SqlClient;
using Microsoft.Data.Sqlite;
using MySqlConnector;
using Npgsql;
using Oracle.ManagedDataAccess.Client;

namespace CrossReportRunner.Database;

public static class DbFactory
{
    public static IDbConnection CreateConnection(DbConnectionConfig config)
    {
        return config.DbType switch
        {
            SupportedDbType.SqlServer => new SqlConnection(
                $"Server={config.ServerName},{config.Port};Database={config.DbName};User Id={config.User};Password={config.Password};TrustServerCertificate=True;Timeout=30;"),
            SupportedDbType.Sqlite => new SqliteConnection(
                $"Data Source={config.DbPath};"),
            SupportedDbType.MySql => new MySqlConnection(
                $"Server={config.ServerName};Port={config.Port};Database={config.DbName};Uid={config.User};Pwd={config.Password};"),
            SupportedDbType.PostgreSql => new NpgsqlConnection(
                $"Host={config.ServerName};Port={config.Port};Database={config.DbName};Username={config.User};Password={config.Password};"),
            SupportedDbType.Oracle => new OracleConnection(
                $"Data Source={config.ServerName}:{config.Port}/{config.DbName};User Id={config.User};Password={config.Password};"),
            _ => throw computerPlatformException(config.DbType)
        };
    }

    private static NotSupportedException computerPlatformException(SupportedDbType type) =>
        new($"Database type {type} is not supported on this platform without native OleDb/Access drivers. Please use SQLite, SQL Server, MySQL, Postgres, or Oracle.");
}
