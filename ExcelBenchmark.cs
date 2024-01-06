using BenchmarkDotNet.Attributes;
using ExcelBattle.Models;
using ExcelBattle.Sut;

namespace ExcelBattle;

[MemoryDiagnoser]
public class ExcelBenchmark
{
    private TemplateData _data = default!;

    [Params(10, 1000, 100000)] public int TotalRow { get; set; }


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
        using var outputStream = new MemoryStream();
        NpoiExcelWriter.Write(_data, outputStream);
    }

    [Benchmark]
    public void UseCloseXML()
    {
        using var outputStream = new MemoryStream();
        CloseXmlExcelWriter.Write(_data, outputStream);
    }

    [Benchmark]
    public void UseMiniExcel()
    {
        using var outputStream = new MemoryStream();
        MiniExcelWriter.Write(_data, outputStream);
    }

    public static void WriteSamples()
    {
        var sample = new ExcelBenchmark
        {
            TotalRow = 100
        };

        sample.GlobalSetup();

        {
            using var outputStream = new MemoryStream();
            NpoiExcelWriter.Write(sample._data, outputStream);
            File.WriteAllBytes("D:\\tmp\\ExportWithCloseXML.xlsx", outputStream.ToArray());
        }


        {
            using var outputStream = new MemoryStream();
            CloseXmlExcelWriter.Write(sample._data, outputStream);
            File.WriteAllBytes("D:\\tmp\\ExportWithNPOI.xlsx", outputStream.ToArray());
        }

        {
            using var outputStream = new MemoryStream();
            MiniExcelWriter.Write(sample._data, outputStream);
            File.WriteAllBytes("D:\\tmp\\ExportWithMiniExcel.xlsx", outputStream.ToArray());
        }
    }
}
