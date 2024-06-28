namespace FastExcelSlim;

[AttributeUsage(AttributeTargets.Class | AttributeTargets.Struct, Inherited = false)]
public sealed class OpenXmlWritableAttribute : Attribute
{
    public string? SheetName { get; set; }

    public bool AutoFilter { get; set; }

    public bool FreezeHeader { get; set; }
}

[AttributeUsage(AttributeTargets.Field | AttributeTargets.Property)]
public sealed class OpenXmlPropertyAttribute : Attribute
{
    public string? ColumnName { get; set; }

    public double Width { get; set; }
}

[AttributeUsage(AttributeTargets.Field | AttributeTargets.Property)]
public sealed class OpenXmlIgnoreAttribute : Attribute;

[AttributeUsage(AttributeTargets.Field | AttributeTargets.Property)]
public sealed class OpenXmlOrderAttribute(int order) : Attribute
{
    public int Order { get; } = order;
}

[AttributeUsage(AttributeTargets.Field | AttributeTargets.Property)]
public sealed class OpenXmlEnumFormatAttribute(string format) : Attribute
{
    public string Format { get; } = format;
}