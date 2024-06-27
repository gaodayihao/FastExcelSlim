namespace FastExcelSlim;

[OpenXmlWritable(SheetName = "Simple")]
internal partial class Simple
{
    public string Name { get; set; }

    public DateTime Birthday { get; set; }

    public Gender Gender { get; set; }
}