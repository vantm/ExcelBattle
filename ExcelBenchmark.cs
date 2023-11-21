using BenchmarkDotNet.Attributes;
using ExcelBattle.Properties;
using IronXL;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace ExcelBattle;

[MemoryDiagnoser]
public class ExcelBenchmark
{
    public Address[] Addresses { get; private set; }
    public Company[] Companies { get; private set; }
    public Contact[] Contacts { get; private set; }
    public Person[] People { get; private set; }
    public Product[] Products { get; private set; }

    [Params(10, 1000, 100000)]
    public int TotalRow { get; set; }

    public readonly byte[] Buffer = new byte[1024 * 1024];

    [GlobalSetup]
    public void GlobalSetup()
    {
        Addresses = TestUtils.GenerateAddresses(TotalRow);
        Companies = TestUtils.GenerateCompanies(TotalRow);
        Contacts = TestUtils.GenerateContacts(TotalRow);
        People = TestUtils.GeneratePeople(TotalRow);
        Products = TestUtils.GenerateProducts(TotalRow);
    }

    [Benchmark]
    public void UseNpoi()
    {
        IWorkbook workbook;

        using (var templateBuffer = new MemoryStream(Resources.ExportTemplate))
        {
            workbook = new XSSFWorkbook(templateBuffer);
        }

        var companySheet = workbook.GetSheet("Company");
        var rowIndex = 0;
        foreach (var company in Companies)
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
        foreach (var address in Addresses)
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
        foreach (var person in People)
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
        foreach (var product in Products)
        {
            var row = productSheet.CreateRow(rowIndex);

            row.CreateCell(0).SetCellValue(product.Name);
            row.CreateCell(1).SetCellValue(product.Price);
            row.CreateCell(2).SetCellValue(product.Description);

            rowIndex++;
        }

        var contactSheet = workbook.GetSheet("Contact");
        rowIndex = 0;
        foreach (var contact in Contacts)
        {
            var row = contactSheet.CreateRow(rowIndex);

            row.CreateCell(0).SetCellValue(contact.FirstName);
            row.CreateCell(1).SetCellValue(contact.LastName);
            row.CreateCell(2).SetCellValue(contact.PhoneNumber);
            row.CreateCell(3).SetCellValue(contact.Email);

            rowIndex++;
        }

        using var mem = new MemoryStream(Buffer);

        workbook.Write(mem, leaveOpen: true);
    }

    [Benchmark]
    public void UseIronXL()
    {
        var workbook = WorkBook.Load(Resources.ExportTemplate);

        var companySheet = workbook.GetWorkSheet("Company");
        var rowIndex = 0;
        foreach (var company in Companies)
        {
            var row = companySheet.InsertRow(rowIndex);

            row.InsertColumn(0).Value = company.Name;
            row.InsertColumn(1).Value = company.HeadquartersStreet;
            row.InsertColumn(2).Value = company.HeadquartersCity;
            row.InsertColumn(3).Value = company.HeadquartersState;
            row.InsertColumn(4).Value = company.HeadquartersZipCode;
            row.InsertColumn(5).Value = company.Revenue;
            row.InsertColumn(6).Value = company.EmployeeCount;

            rowIndex++;
        }

        var addressSheet = workbook.GetWorkSheet("Address");
        rowIndex = 0;
        foreach (var address in Addresses)
        {
            var row = addressSheet.InsertRow(rowIndex);

            row.InsertColumn(0).Value = address.City;
            row.InsertColumn(1).Value = address.Street;
            row.InsertColumn(2).Value = address.State;
            row.InsertColumn(3).Value = address.ZipCode;

            rowIndex++;
        }

        var peopleSheet = workbook.GetWorkSheet("People");
        rowIndex = 0;
        foreach (var person in People)
        {
            var row = peopleSheet.InsertRow(rowIndex);

            row.InsertColumn(0).Value = person.FirstName;
            row.InsertColumn(1).Value = person.LastName;
            row.InsertColumn(2).Value = person.Age;
            row.InsertColumn(3).Value = person.Address;

            rowIndex++;
        }

        var productSheet = workbook.GetWorkSheet("Product");
        rowIndex = 0;
        foreach (var product in Products)
        {
            var row = productSheet.InsertRow(rowIndex);

            row.InsertColumn(0).Value = product.Name;
            row.InsertColumn(1).Value = product.Price;
            row.InsertColumn(2).Value = product.Description;

            rowIndex++;
        }

        var contactSheet = workbook.GetWorkSheet("Contact");
        rowIndex = 0;
        foreach (var contact in Contacts)
        {
            var row = contactSheet.InsertRow(rowIndex);

            row.InsertColumn(0).Value = contact.FirstName;
            row.InsertColumn(1).Value = contact.LastName;
            row.InsertColumn(2).Value = contact.PhoneNumber;
            row.InsertColumn(3).Value = contact.Email;

            rowIndex++;
        }

        using var mem = new MemoryStream(Buffer);

        using var stream = workbook.ToStream();

        stream.CopyTo(mem);
    }

    public static void MakeSample()
    {
        var sample = new ExcelBenchmark();

        sample.TotalRow = 1000;

        sample.GlobalSetup();

        sample.UseNpoi();
        File.WriteAllBytes("D:\\tmp\\ExportWithNPOI.xlsx", sample.Buffer);

        Array.Clear(sample.Buffer, 0, sample.Buffer.Length);
        sample.UseIronXL();
        File.WriteAllBytes("D:\\tmp\\ExportWithIronXL.xlsx", sample.Buffer);
    }
}