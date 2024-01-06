using Bogus;
using ExcelBattle.Models;
using Person = ExcelBattle.Models.Person;

namespace ExcelBattle;

public static class TestUtils
{
    private static readonly Faker _faker = new();

    public static Company GenerateCompany()
    {
        return new Company(
            _faker.Company.CompanyName(),
            _faker.Address.StreetAddress(),
            _faker.Address.City(),
            _faker.Address.State(),
            _faker.Address.ZipCode(),
            (double)_faker.Random.Decimal(1000, 1000000),
            _faker.Random.Int(10, 10000)
        );
    }

    public static Company[] GenerateCompanies(int count)
    {
        var companies = new List<Company>();
        for (var i = 0; i < count; i++)
        {
            companies.Add(GenerateCompany());
        }
        return companies.ToArray();
    }

    public static Product GenerateProduct()
    {
        return new Product(
            _faker.Commerce.ProductName(),
            _faker.Commerce.ProductDescription(),
            (double)_faker.Finance.Amount(10, 1000)
        );
    }

    public static Product[] GenerateProducts(int count)
    {
        var products = new List<Product>();
        for (var i = 0; i < count; i++)
        {
            products.Add(GenerateProduct());
        }
        return products.ToArray();
    }

    public static Contact GenerateContact()
    {
        return new Contact(
            _faker.Name.FirstName(),
            _faker.Name.LastName(),
            _faker.Internet.Email(),
            _faker.Phone.PhoneNumber()
        );
    }

    public static Contact[] GenerateContacts(int count)
    {
        var contacts = new List<Contact>();
        for (var i = 0; i < count; i++)
        {
            contacts.Add(GenerateContact());
        }
        return contacts.ToArray();
    }

    public static Address GenerateAddress()
    {
        return new Address(
            _faker.Address.StreetAddress(),
            _faker.Address.City(),
            _faker.Address.State(),
            _faker.Address.ZipCode()
        );
    }

    public static Address[] GenerateAddresses(int count)
    {
        var addresses = new List<Address>();
        for (var i = 0; i < count; i++)
        {
            addresses.Add(GenerateAddress());
        }
        return addresses.ToArray();
    }

    public static Person GeneratePerson()
    {
        return new Person(
            _faker.Name.FirstName(),
            _faker.Name.LastName(),
            _faker.Random.Int(18, 80),
            _faker.Address.StreetAddress()
        );
    }

    public static Person[] GeneratePeople(int count)
    {
        var people = new List<Person>();
        for (var i = 0; i < count; i++)
        {
            people.Add(GeneratePerson());
        }
        return people.ToArray();
    }
}
