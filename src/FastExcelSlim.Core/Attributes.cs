namespace FastExcelSlim;

[AttributeUsage(AttributeTargets.Class | AttributeTargets.Struct, Inherited = false)]
public sealed class OpenXmlWritableAttribute : Attribute
{
    public OpenXmlWritableAttribute()
    {

    }

    public OpenXmlWritableAttribute(string? sheetName)
    {
        SheetName = sheetName;
    }

    public string? SheetName { get; }
}

[AttributeUsage(AttributeTargets.Field | AttributeTargets.Property)]
public sealed class OpenXmlPropertyAttribute : Attribute
{
    public OpenXmlPropertyAttribute()
    {
    }

    public OpenXmlPropertyAttribute(string columnName)
    {
        ColumnName = columnName;
    }

    public OpenXmlPropertyAttribute(string columnName, int width)
    {
        ColumnName = columnName;
        Width = width;
    }

    public OpenXmlPropertyAttribute(int width)
    {
        Width = width;
    }

    public string? ColumnName { get; }

    public int? Width { get; }
}

[AttributeUsage(AttributeTargets.Field | AttributeTargets.Property)]
public sealed class OpenXmlIgnoreAttribute : Attribute;

[AttributeUsage(AttributeTargets.Field | AttributeTargets.Property)]
public sealed class OpenXmlOrderAttribute(int order) : Attribute
{
    public int Order { get; } = order;
}