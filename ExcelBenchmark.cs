using BenchmarkDotNet.Attributes;
using ExcelBattle.Models;
using ExcelBattle.Sut;

namespace ExcelBattle;

[MemoryDiagnoser]
public class ExcelBenchmark
{
    private TemplateData _data = default!;

    [Params(10, 1000, 100000)] public int TotalRow { get; set; }

    private const string TemplateDir = "D:\\tmp\\";

    private static string GenerateOutputFileName() => $"{TemplateDir}\\{Guid.NewGuid()}.xlsx";

    [GlobalSetup]
    public void GlobalSetup()
    {
        _data = new TemplateData
        {
            Addresses = TestUtils.GenerateAddresses(TotalRow),
            Companies = TestUtils.GenerateCompanies(TotalRow),
            Contacts = TestUtils.GenerateContacts(TotalRow),
            People = TestUtils.GeneratePeople(TotalRow),
            Products = TestUtils.GenerateProducts(TotalRow)
        };
    }

    [Benchmark]
    public void UseNpoi()
    {
        var fileName = GenerateOutputFileName();
        NpoiExcelWriter.Write(_data, fileName);
    }

    [Benchmark]
    public void UseCloseXML()
    {
        var fileName = GenerateOutputFileName();
        CloseXmlExcelWriter.Write(_data, fileName);
    }

    [Benchmark]
    public void UseMiniExcel()
    {
        var fileName = GenerateOutputFileName();
        MiniExcelWriter.Write(_data, fileName);
    }

    public static void WriteSamples()
    {
        var sample = new ExcelBenchmark
        {
            TotalRow = 100
        };

        sample.GlobalSetup();

        NpoiExcelWriter.Write(sample._data, $"{TemplateDir}\\ExportWithNpoi.xlsx");
        CloseXmlExcelWriter.Write(sample._data, $"{TemplateDir}\\ExportWithCloseXML.xlsx");
        MiniExcelWriter.Write(sample._data, $"{TemplateDir}\\ExportWithMiniExcel.xlsx");
    }
}
