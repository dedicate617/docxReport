using CrossReportRunner.Parsers;
using CrossReportRunner.Models;
using Xunit;

namespace CrossReportRunner.Tests;

public class DsDefParserTests
{
    private const string SampleDsdef = @"<Report>
  <Tables>
    <Table>
      <Name>LYO1_LYO_Header</Name>
      <DataSource>LYO1</DataSource>
      <Sql>SELECT [BatchNo] FROM [LYO1_LYO_Header] WHERE [BatchNo] = '[%BatchNo%]'</Sql>
      <EnableSampling>False</EnableSampling>
      <Fields>
        <Field><Name>BatchNo</Name><FieldType>0</FieldType></Field>
      </Fields>
    </Table>
    <Table>
      <Name>LYO1_LYO_ProcessValue</Name>
      <DataSource>LYO1</DataSource>
      <Sql>SELECT [TE201],[RecordTime] FROM [LYO1_LYO_ProcessValue] WHERE [BatchNo] = '[%BatchNo%]'</Sql>
      <EnableSampling>True</EnableSampling>
      <SamplingDistance>60</SamplingDistance>
      <SamplingField>RecordTime</SamplingField>
      <SamplingType>0</SamplingType>
      <TimeUnit>5</TimeUnit>
    </Table>
  </Tables>
</Report>";

    [Fact]
    public void Parse_ValidXml_ReturnsPopulatedDefinition()
    {
        var result = DsDefParser.Parse(SampleDsdef);
        Assert.NotNull(result);
        Assert.Equal(2, result.Tables.Count);
        
        var header = result.Tables[0];
        Assert.Equal("LYO1_LYO_Header", header.Name);
        Assert.Equal("LYO1", header.DataSource);
        Assert.False(header.EnableSampling);
        Assert.Single(header.Fields);

        var process = result.Tables[1];
        Assert.Equal("LYO1_LYO_ProcessValue", process.Name);
        Assert.True(process.EnableSampling);
        Assert.Equal(60, process.SamplingDistance);
        Assert.Equal("RecordTime", process.SamplingField);
        Assert.Equal(TimeUnit.Second, process.TimeUnit);
    }
}
