using FastExcelSlim.Extensions;
using Utf8StringInterpolation;

namespace FastExcelSlim.OpenXml;

public class CellStyle : IdXmlElement
{
    private readonly bool _writeId;
    public CellAlignment? Alignment { get; private set; }

    public CellProtection? Protection { get; private set; }

    public DataFormat? DataFormat { get; set; }

    public Font? Font { get; set; }

    public Fill? Fill { get; set; }

    public Border? Border { get; set; }

    internal int DataFormatId => DataFormat?.Id ?? 0;
    internal int FontId => Font?.Id ?? 0;
    internal int FillId => Fill?.Id ?? 0;
    internal int BorderId => Border?.Id ?? 0;

    internal bool ApplyDataFormat => DataFormat != null;
    internal bool ApplyFont => Font != null;
    internal bool ApplyFill => Fill != null;
    internal bool ApplyBorder => Border != null;
    internal bool ApplyAlignment => Alignment != null;
    internal bool ApplyProtection => Protection != null;

    public CellAlignment AddAlignment()
    {
        return Alignment ??= new CellAlignment();
    }

    public CellProtection AddProtection()
    {
        return Protection ??= new CellProtection();
    }

    internal CellStyle(int id, bool writeId) : base(id)
    {
        _writeId = writeId;
    }

    internal override void Write<TBufferWriter>(scoped ref Utf8StringWriter<TBufferWriter> writer)
    {
        writer.AppendLiteral("<xf");
        writer.WriteAttribute("numFmtId", DataFormatId, true);
        writer.WriteAttribute("fontId", FontId, true);
        writer.WriteAttribute("fillId", FillId, true);
        writer.WriteAttribute("borderId", BorderId, true);
        if (_writeId) writer.WriteAttribute("xfId", 0, true);
        writer.WriteAttribute("applyNumberFormat", ApplyDataFormat);
        writer.WriteAttribute("applyFont", ApplyFont);
        writer.WriteAttribute("applyFill", ApplyFill);
        writer.WriteAttribute("applyBorder", ApplyBorder);
        writer.WriteAttribute("applyAlignment", ApplyAlignment);
        writer.WriteAttribute("applyProtection", ApplyProtection);

        if (Alignment == null && Protection == null)
        {
            writer.AppendLiteral("/>");
        }
        else
        {
            writer.AppendLiteral(">");
            Alignment?.Write(ref writer);
            Protection?.Write(ref writer);
            writer.AppendLiteral("</xf>");
        }
    }
}

public class CellStyleSheet : XmlElementCollection<CellStyle>
{
    internal void Initialize()
    {
        Add(new CellStyle(Count, false));
    }

    internal override string Name => "cellStyleXfs";
}

public class CellStyles : XmlElementCollection<CellStyle>
{
    internal CellStyles() { }

    internal override string Name => "cellXfs";

    public CellStyle CreateCellStyle()
    {
        var style = new CellStyle(Count, true);
        Add(style);
        return style;
    }
}