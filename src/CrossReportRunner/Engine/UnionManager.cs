using CrossReportRunner.Models;

namespace CrossReportRunner.Engine;

public static class UnionManager
{
    public static TableData MergeAndSort(TableData baseTable, TableData unionTable, string secondarySort)
    {
        var mergedColumns = new List<string>(baseTable.Columns);
        foreach (var col in unionTable.Columns)
        {
            if (!mergedColumns.Contains(col, StringComparer.OrdinalIgnoreCase))
                mergedColumns.Add(col);
        }

        var result = new TableData { Columns = mergedColumns };

        foreach (var row in baseTable.Rows)
        {
            result.Rows.Add(AlignRow(row, baseTable.Columns, mergedColumns));
        }

        foreach (var row in unionTable.Rows)
        {
            result.Rows.Add(AlignRow(row, unionTable.Columns, mergedColumns));
        }

        if (!string.IsNullOrEmpty(secondarySort))
        {
            var parts = secondarySort.Trim().Split(' ', StringSplitOptions.RemoveEmptyEntries);
            string sortCol = parts[0];
            bool descending = parts.Length > 1 && parts[1].Equals("DESC", StringComparison.OrdinalIgnoreCase);

            int sortColIdx = mergedColumns.FindIndex(c => c.Equals(sortCol, StringComparison.OrdinalIgnoreCase));
            if (sortColIdx >= 0)
            {
                result.Rows.Sort((a, b) =>
                {
                    var va = a[sortColIdx]?.ToString() ?? string.Empty;
                    var vb = b[sortColIdx]?.ToString() ?? string.Empty;
                    int cmp = string.Compare(va, vb, StringComparison.Ordinal);
                    return descending ? -cmp : cmp;
                });
            }
        }

        return result;
    }

    private static List<object?> AlignRow(List<object?> sourceRow, List<string> sourceCols, List<string> targetCols)
    {
        var row = new List<object?>(targetCols.Count);
        for (int i = 0; i < targetCols.Count; i++)
        {
            int sIdx = sourceCols.FindIndex(c => c.Equals(targetCols[i], StringComparison.OrdinalIgnoreCase));
            row.Add(sIdx >= 0 && sIdx < sourceRow.Count ? sourceRow[sIdx] : null);
        }
        return row;
    }
}
