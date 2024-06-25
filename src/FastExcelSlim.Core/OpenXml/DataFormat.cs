using FastExcelSlim.Extensions;
using Utf8StringInterpolation;

namespace FastExcelSlim.OpenXml;

public class DataFormat : IdXmlElement
{
    internal DataFormat(int id, string formatCode) : base(id)
    {
        FormatCode = formatCode;
    }

    public string FormatCode { get; }

    public bool IsBuiltinFormat => Id < BuiltinFormats.FirstUserDefinedFormatIndex;

    internal override void Write<TBufferWriter>(scoped ref Utf8StringWriter<TBufferWriter> writer)
    {
        if (IsBuiltinFormat) return;

        writer.AppendFormat($"<numFmt");
        writer.WriteAttribute("numFmtId", Id, true);
        writer.WriteAttribute("formatCode", FormatCode);
        writer.AppendLiteral("/>");
    }
}

public class DataFormats : XmlElementCollection<DataFormat>
{
    internal override string Name => "numFmts";

    public DataFormat GetDateFormat(string format)
    {
        var id = BuiltinFormats.GetBuiltinFormat(format);
        if (id == -1)
        {
            var exists = InnerElements.FirstOrDefault(f => f.FormatCode == format);
            if (exists != null) return exists;

            id = BuiltinFormats.FirstUserDefinedFormatIndex + Count;
            var dataFormat = new DataFormat(id, format);
            Add(dataFormat);
            return dataFormat;
        }

        return new DataFormat(id, format);
    }

    protected override void Add(DataFormat item)
    {
        if (item.IsBuiltinFormat) return;
        base.Add(item);
    }
}