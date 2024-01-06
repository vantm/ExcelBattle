namespace ExcelBattle.Models;

public sealed class TemplateData
{
    public required Address[] Addresses { get; init; }
    public required Company[] Companies { get; init; }
    public required Contact[] Contacts { get; init; }
    public required Person[] People { get; init; }
    public required Product[] Products { get; init; }
}
