using ExcelBattle.Properties;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace ExcelBattle;

public static class NpoiExcelWriter
{
    public static void Write(TemplateData data, Stream outputStream)
    {
        IWorkbook workbook;

        using (var templateBuffer = new MemoryStream(Resources.ExportTemplate))
        {
            workbook = new XSSFWorkbook(templateBuffer);
        }

        var companySheet = workbook.GetSheet("Company");
        var rowIndex = 0;
        foreach (var company in data.Companies)
        {
            var row = companySheet.CreateRow(rowIndex);

            row.CreateCell(0).SetCellValue(company.Name);
            row.CreateCell(1).SetCellValue(company.HeadquartersStreet);
            row.CreateCell(2).SetCellValue(company.HeadquartersCity);
            row.CreateCell(3).SetCellValue(company.HeadquartersState);
            row.CreateCell(4).SetCellValue(company.HeadquartersZipCode);
            row.CreateCell(5).SetCellValue(company.Revenue);
            row.CreateCell(6).SetCellValue(company.EmployeeCount);

            rowIndex++;
        }

        var addressSheet = workbook.GetSheet("Address");
        rowIndex = 0;
        foreach (var address in data.Addresses)
        {
            var row = addressSheet.CreateRow(rowIndex);

            row.CreateCell(0).SetCellValue(address.City);
            row.CreateCell(1).SetCellValue(address.Street);
            row.CreateCell(2).SetCellValue(address.State);
            row.CreateCell(3).SetCellValue(address.ZipCode);

            rowIndex++;
        }

        var peopleSheet = workbook.GetSheet("People");
        rowIndex = 0;
        foreach (var person in data.People)
        {
            var row = peopleSheet.CreateRow(rowIndex);

            row.CreateCell(0).SetCellValue(person.FirstName);
            row.CreateCell(1).SetCellValue(person.LastName);
            row.CreateCell(2).SetCellValue(person.Age);
            row.CreateCell(3).SetCellValue(person.Address);

            rowIndex++;
        }

        var productSheet = workbook.GetSheet("Product");
        rowIndex = 0;
        foreach (var product in data.Products)
        {
            var row = productSheet.CreateRow(rowIndex);

            row.CreateCell(0).SetCellValue(product.Name);
            row.CreateCell(1).SetCellValue(product.Price);
            row.CreateCell(2).SetCellValue(product.Description);

            rowIndex++;
        }

        var contactSheet = workbook.GetSheet("Contact");
        rowIndex = 0;
        foreach (var contact in data.Contacts)
        {
            var row = contactSheet.CreateRow(rowIndex);

            row.CreateCell(0).SetCellValue(contact.FirstName);
            row.CreateCell(1).SetCellValue(contact.LastName);
            row.CreateCell(2).SetCellValue(contact.PhoneNumber);
            row.CreateCell(3).SetCellValue(contact.Email);

            rowIndex++;
        }

        workbook.Write(outputStream, leaveOpen: true);
    }
}
