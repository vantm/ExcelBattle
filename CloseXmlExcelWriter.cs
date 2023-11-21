using ClosedXML.Excel;
using ExcelBattle.Properties;

namespace ExcelBattle;

public static class CloseXmlExcelWriter
{
    public static void Write(TemplateData data, Stream outputStream)
    {
        using var templateBuffer = new MemoryStream(Resources.ExportTemplate);
        using IXLWorkbook workbook = new XLWorkbook(templateBuffer);

        workbook.TryGetWorksheet("Company", out var companySheet);

        var rowIndex = 1;
        foreach (var company in data.Companies)
        {
            var row = companySheet.Row(rowIndex);

            row.Cell(1).Value = company.Name;
            row.Cell(2).Value = company.HeadquartersStreet;
            row.Cell(3).Value = company.HeadquartersCity;
            row.Cell(4).Value = company.HeadquartersState;
            row.Cell(5).Value = company.HeadquartersZipCode;
            row.Cell(6).Value = company.Revenue;
            row.Cell(7).Value = company.EmployeeCount;

            rowIndex++;
        }

        workbook.TryGetWorksheet("Address", out var addressSheet);
        rowIndex = 1;
        foreach (var address in data.Addresses)
        {
            var row = addressSheet.Row(rowIndex);

            row.Cell(1).Value = address.City;
            row.Cell(2).Value = address.Street;
            row.Cell(3).Value = address.State;
            row.Cell(4).Value = address.ZipCode;

            rowIndex++;
        }

        workbook.TryGetWorksheet("People", out var peopleSheet);
        rowIndex = 1;
        foreach (var person in data.People)
        {
            var row = peopleSheet.Row(rowIndex);

            row.Cell(1).Value = person.FirstName;
            row.Cell(2).Value = person.LastName;
            row.Cell(3).Value = person.Age;
            row.Cell(4).Value = person.Address;

            rowIndex++;
        }

        workbook.TryGetWorksheet("Product", out var productSheet);
        rowIndex = 1;
        foreach (var product in data.Products)
        {
            var row = productSheet.Row(rowIndex);

            row.Cell(1).Value = product.Name;
            row.Cell(2).Value = product.Price;
            row.Cell(3).Value = product.Description;

            rowIndex++;
        }

        workbook.TryGetWorksheet("Contact", out var contactSheet);
        rowIndex = 1;
        foreach (var contact in data.Contacts)
        {
            var row = contactSheet.Row(rowIndex);

            row.Cell(1).Value = contact.FirstName;
            row.Cell(2).Value = contact.LastName;
            row.Cell(3).Value = contact.PhoneNumber;
            row.Cell(4).Value = contact.Email;

            rowIndex++;
        }

        workbook.SaveAs(outputStream, false);
    }
}
