using ExcelBattle.Models;
using ExcelBattle.Properties;
using MiniExcelLibs;

namespace ExcelBattle.Sut;

public static class MiniExcelWriter
{
    public static void Write(TemplateData data, Stream outputStream)
    {
        var templateData = new Dictionary<string, object>
        {
            ["Companies"] = data.Companies,
            ["Addresses"] = data.Addresses,
            ["Contacts"] = data.Contacts,
            ["Products"] = data.Products,
            ["People"] = data.People
        };

        outputStream.SaveAsByTemplate(Resources.MiniExcelExportTemplate, templateData);
    }
}
