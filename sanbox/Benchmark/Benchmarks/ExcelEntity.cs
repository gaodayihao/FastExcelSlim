namespace FastExcelSlim.Benchmarks;

[OpenXmlWritable]
internal partial class ExcelEntity
{
    public int Id { get; set; }

    public string? Name { get; set; }

    public Guid GUIId { get; set; }

    public int Age { get; set; }

    public DateTime Birthday { get; set; }

    public string? Class { get; set; }

    public string? Address { get; set; }

    public double Deposit { get; set; }

    public long Friends { get; set; }

    public bool IsOnline { get; set; }

    public DateOnly LastOnline { get; set; }
}