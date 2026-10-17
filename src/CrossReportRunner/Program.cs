using System.Data;
using CrossReportRunner.Database;
using CrossReportRunner.Engine;
using CrossReportRunner.Models;
using CrossReportRunner.Parsers;
using CrossReportRunner.Serialization;

namespace CrossReportRunner;

public class Program
{
    public static int Main(string[] args)
    {
        string dsdefPath = string.Empty;
        string dspara = string.Empty;
        string vpara = string.Empty;
        string outputMode = "pipe-gzip";
        string exportJson = string.Empty;

        for (int i = 0; i < args.Length; i++)
        {
            var arg = args[i];
            if (arg.StartsWith("--dsdef=")) dsdefPath = arg.Substring(8).Trim('"');
            else if (arg.StartsWith("-s=") || arg.Equals("-s")) dsdefPath = (arg.Length > 2 ? arg.Substring(3) : args[++i]).Trim('"');
            else if (arg.StartsWith("--dspara=")) dspara = arg.Substring(9).Trim('"');
            else if (arg.StartsWith("-a=") || arg.Equals("-a")) dspara = (arg.Length > 2 ? arg.Substring(3) : args[++i]).Trim('"');
            else if (arg.StartsWith("--vpara=")) vpara = arg.Substring(8).Trim('"');
            else if (arg.StartsWith("-r=") || arg.Equals("-r")) vpara = (arg.Length > 2 ? arg.Substring(3) : args[++i]).Trim('"');
            else if (arg.StartsWith("--output-mode=")) outputMode = arg.Substring(14).Trim('"');
            else if (arg.StartsWith("--export-json=")) exportJson = arg.Substring(14).Trim('"');
        }

        if (string.IsNullOrEmpty(dsdefPath) || !File.Exists(dsdefPath))
        {
            Console.Error.WriteLine($"[Error] dsdef file invalid: {dsdefPath}");
            return 2;
        }

        try
        {
            var reportDef = DsDefParser.ParseFile(dsdefPath);
            var connConfig = DbConnectionConfig.Parse(dspara);

            var variables = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            if (!string.IsNullOrEmpty(vpara))
            {
                foreach (var item in vpara.Split(';', StringSplitOptions.RemoveEmptyEntries))
                {
                    var kv = item.Split('=', 2);
                    if (kv.Length == 2) variables[kv[0].Trim()] = kv[1].Trim();
                }
            }

            var payload = new CompactDataPayload { Variables = variables };

            using var connection = DbFactory.CreateConnection(connConfig);
            connection.Open();

            // Execute Queries
            foreach (var tableDef in reportDef.Tables)
            {
                var tableData = QueryTable(connection, tableDef, variables);
                if (tableDef.EnableSampling)
                {
                    tableData = SamplingEngine.ApplySampling(tableData, tableDef);
                }
                payload.Tables[tableDef.Name] = tableData;
            }

            // Handle Unions
            foreach (var tableDef in reportDef.Tables)
            {
                if (!string.IsNullOrEmpty(tableDef.Union) && payload.Tables.ContainsKey(tableDef.Union))
                {
                    var baseTable = payload.Tables[tableDef.Union];
                    var unionTable = payload.Tables[tableDef.Name];
                    payload.Tables[tableDef.Union] = UnionManager.MergeAndSort(baseTable, unionTable, tableDef.SecondarySort);
                }
            }

            if (!string.IsNullOrEmpty(exportJson))
            {
                StreamSerializer.WriteFilePayload(payload, exportJson);
            }

            if (outputMode == "pipe-gzip")
            {
                using var stdout = Console.OpenStandardOutput();
                StreamSerializer.WriteGzipPayload(payload, stdout);
            }

            return 0;
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"[Error] Execution failed: {ex.Message}");
            return 4;
        }
    }

    private static TableData QueryTable(IDbConnection conn, DsTableDefinition def, Dictionary<string, string> vars)
    {
        var resolvedSql = SqlParameterResolver.ResolveVariables(def.Sql, vars);
        using var cmd = conn.CreateCommand();
        cmd.CommandText = resolvedSql;
        using var reader = cmd.ExecuteReader();

        var data = new TableData();
        for (int i = 0; i < reader.FieldCount; i++)
            data.Columns.Add(reader.GetName(i));

        while (reader.Read())
        {
            var row = new List<object?>(reader.FieldCount);
            for (int i = 0; i < reader.FieldCount; i++)
            {
                row.Add(reader.IsDBNull(i) ? null : reader.GetValue(i));
            }
            data.Rows.Add(row);
        }

        return data;
    }
}
