namespace ExcelBattle.Models;

public record Company(
    string Name,
    string HeadquartersStreet,
    string HeadquartersCity,
    string HeadquartersState,
    string HeadquartersZipCode,
    double Revenue,
    int EmployeeCount
);
