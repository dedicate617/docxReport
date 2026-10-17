using System.Xml.Linq;
using CrossReportRunner.Models;

namespace CrossReportRunner.Parsers;

public static class DsDefParser
{
    public static DsReportDefinition ParseFile(string filePath)
    {
        if (!File.Exists(filePath))
            throw new FileNotFoundException($"dsdef file not found: {filePath}");

        var content = File.ReadAllText(filePath);
        return Parse(content);
    }

    public static DsReportDefinition Parse(string xmlContent)
    {
        var doc = XDocument.Parse(xmlContent);
        var reportDef = new DsReportDefinition();

        var tableElements = doc.Descendants("Table");
        foreach (var tblElem in tableElements)
        {
            // Only parse top-level Table elements here, children will be handled hierarchically
            if (tblElem.Parent?.Name == "Tables" && tblElem.Parent?.Parent?.Name == "Report")
            {
                reportDef.Tables.Add(ParseTable(tblElem));
            }
        }

        return reportDef;
    }

    private static DsTableDefinition ParseTable(XElement element)
    {
        var table = new DsTableDefinition
        {
            Name = element.Element("Name")?.Value.Trim() ?? string.Empty,
            DataSource = element.Element("DataSource")?.Value.Trim() ?? string.Empty,
            Sql = element.Element("Sql")?.Value.Trim() ?? string.Empty,
            Union = element.Element("Union")?.Value.Trim() ?? string.Empty,
            RelateFields = element.Element("RelateFields")?.Value.Trim() ?? string.Empty,
            SecondarySort = element.Element("SecondarySort")?.Value.Trim() ?? string.Empty,
            SamplingField = element.Element("SamplingField")?.Value.Trim() ?? string.Empty,
        };

        if (bool.TryParse(element.Element("EnableSampling")?.Value, out var enableSampling))
            table.EnableSampling = enableSampling;

        if (int.TryParse(element.Element("SamplingDistance")?.Value, out var samplingDistance))
            table.SamplingDistance = samplingDistance;

        if (int.TryParse(element.Element("SamplingType")?.Value, out var samplingType))
            table.SamplingType = (SamplingType)samplingType;

        if (int.TryParse(element.Element("TimeUnit")?.Value, out var timeUnit))
            table.TimeUnit = (TimeUnit)timeUnit;

        var fieldsElem = element.Element("Fields");
        if (fieldsElem != null)
        {
            foreach (var fld in fieldsElem.Elements("Field"))
            {
                table.Fields.Add(new DsFieldDefinition
                {
                    Name = fld.Element("Name")?.Value.Trim() ?? string.Empty,
                    FieldType = fld.Element("FieldType")?.Value.Trim() ?? "0"
                });
            }
        }

        var subTablesElem = element.Element("Tables");
        if (subTablesElem != null)
        {
            foreach (var subTbl in subTablesElem.Elements("Table"))
            {
                table.SubTables.Add(ParseTable(subTbl));
            }
        }

        return table;
    }
}
