namespace FastExcelSlim;

[AttributeUsage(AttributeTargets.Class | AttributeTargets.Struct, Inherited = false)]
public sealed class OpenXmlWritableAttribute : Attribute
{
    public string? SheetName { get; set; }
}

[AttributeUsage(AttributeTargets.Field | AttributeTargets.Property)]
public sealed class OpenXmlPropertyAttribute : Attribute
{
    public string? ColumnName { get; set; }

    public int Width { get; set; } = 15;
}


[AttributeUsage(AttributeTargets.Field | AttributeTargets.Property)]
public sealed class OpenXmlIgnoreAttribute : Attribute;