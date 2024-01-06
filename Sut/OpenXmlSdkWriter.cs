using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ExcelBattle.Models;
using ExcelBattle.Properties;

namespace ExcelBattle.Sut;

public static class OpenXmlSdkWriter
{
    public static void Write(TemplateData data, string path)
    {
        // copy files
        using var stream = File.Open(path, FileMode.Create, FileAccess.ReadWrite);

        stream.Write(Resources.ExportTemplate);

        stream.Seek(0, SeekOrigin.Begin);

        var doc = SpreadsheetDocument.Open(stream, true);

        WriteToSheet("Company", data.Companies, (row, company) =>
        {
            row.InsertAt(GenerateNewCell(company.Name), 0);
            row.InsertAt(GenerateNewCell(company.HeadquartersStreet), 1);
            row.InsertAt(GenerateNewCell(company.HeadquartersCity), 2);
            row.InsertAt(GenerateNewCell(company.HeadquartersState), 3);
            row.InsertAt(GenerateNewCell(company.HeadquartersZipCode), 4);
            row.InsertAt(GenerateNewCell(company.Revenue), 5);
            row.InsertAt(GenerateNewCell(company.EmployeeCount), 6);
        });

        WriteToSheet("Address", data.Addresses, (row, address) =>
        {
            row.InsertAt(GenerateNewCell(address.City), 0);
            row.InsertAt(GenerateNewCell(address.Street), 1);
            row.InsertAt(GenerateNewCell(address.State), 2);
            row.InsertAt(GenerateNewCell(address.ZipCode), 3);
        });

        WriteToSheet("People", data.People, (row, person) =>
        {
            row.InsertAt(GenerateNewCell(person.FirstName), 0);
            row.InsertAt(GenerateNewCell(person.LastName), 1);
            row.InsertAt(GenerateNewCell(person.Age), 2);
            row.InsertAt(GenerateNewCell(person.Address), 3);
        });

        WriteToSheet("Product", data.Products, (row, product) =>
        {
            row.InsertAt(GenerateNewCell(product.Name), 0);
            row.InsertAt(GenerateNewCell(product.Description), 1);
            row.InsertAt(GenerateNewCell(product.Price), 2);
        });

        WriteToSheet("Contact", data.Contacts, (row, contact) =>
        {
            row.InsertAt(GenerateNewCell(contact.FirstName), 0);
            row.InsertAt(GenerateNewCell(contact.LastName), 1);
            row.InsertAt(GenerateNewCell(contact.Email), 2);
            row.InsertAt(GenerateNewCell(contact.PhoneNumber), 3);
        });

        doc.Save();

        void WriteToSheet<T>(string sheetName, IEnumerable<T> rowData, Action<Row, T> writeRowAction)
        {
            var sheet = FindWorksheetPartByName(doc, sheetName)!;
            var sheetData = sheet.Worksheet.GetFirstChild<SheetData>()!;
            var index = 1u;
            foreach (var item in rowData)
            {
                var row = new Row { RowIndex = index };

                writeRowAction(row, item);

                sheetData.Append(row);
                index++;
            }
        }
    }

    private static Cell GenerateNewCell(string value) => new() { CellValue = new(value), DataType = new(CellValues.String) };
    private static Cell GenerateNewCell(int number) => new() { CellValue = new(number), DataType = new(CellValues.Number) };
    private static Cell GenerateNewCell(double number) => new() { CellValue = new(number), DataType = new(CellValues.Number) };

    private static WorksheetPart? FindWorksheetPartByName(SpreadsheetDocument document, string sheetName)
    {
        var sheets = document
            .WorkbookPart
            ?.Workbook
            .GetFirstChild<Sheets>()
            ?.Elements<Sheet>().Where(s => s.Name == sheetName)
            .ToArray();

        if (sheets == null || sheets.Length == 0)
        {
            // The specified worksheet does not exist.
            return null;
        }

        var relationshipId = sheets[0].Id!.Value!;
        var worksheetPart = (WorksheetPart?)document.WorkbookPart!.GetPartById(relationshipId);

        return worksheetPart;
    }
}
