namespace FastExcelSlim;

[OpenXmlWritable(SheetName = "Demo")]
public partial class DemoEntity
{
    [OpenXmlOrder(0)]
    public int Id { get; set; }

    [OpenXmlOrder(2)]
    public string? Name { get; set; }

    [OpenXmlOrder(1)]
    [OpenXmlProperty(ColumnName = "Student Number", Width = 50)]
    public string? Number { get; set; }

    [OpenXmlOrder(3)]
    [OpenXmlProperty(ColumnName = "Gender", Width = 20)]
    [OpenXmlEnumFormat("G")]
    public Gender Gender { get; set; }

    [OpenXmlOrder(4)]
    [OpenXmlProperty(ColumnName = "Gender Value", Width = 20)]
    [OpenXmlEnumFormat("D")]
    private Gender NumberFormatGender => Gender;

    [OpenXmlOrder(5)]
    [OpenXmlProperty(ColumnName = "Age")]
    private int? _age = 5;

    [OpenXmlOrder(6)]
    [OpenXmlProperty(Width = 35)]
    public DateTime BirthDay { get; set; }

    [OpenXmlIgnore]
    public DateTime LastOnline { get; set; }
}

public enum Gender
{
    Male = 1,
    Female,
    Other
}
