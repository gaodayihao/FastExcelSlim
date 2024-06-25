using FastExcelSlim.Extensions;
using Utf8StringInterpolation;

namespace FastExcelSlim.OpenXml;

public class CellAlignment : XmlElement
{
    internal CellAlignment() { }

    private long _textRotation;

    public CellHorizontalAlignment? Horizontal { get; set; }

    public CellVerticalAlignment? Vertical { get; set; }

    public long TextRotation
    {
        get => _textRotation;
        set
        {
            var rotation = value;
            if (rotation is < 0 and >= -90)
            {
                rotation = 90 + (-1 * rotation);
            }
            _textRotation = rotation;
        }
    }

    public long Indent { get; set; }

    public bool WrapText { get; set; }

    public bool ShrinkToFit { get; set; }

    internal override void Write<TBufferWriter>(scoped ref Utf8StringWriter<TBufferWriter> writer)
    {
        writer.AppendLiteral("<alignment");
        if (Horizontal != null && Horizontal != CellHorizontalAlignment.General)
        {
            writer.WriteAttribute("horizontal", Horizontal.Value);
        }
        if (Vertical != null && Vertical != CellVerticalAlignment.Bottom)
        {
            writer.WriteAttribute("vertical", Vertical.Value);
        }
        writer.WriteAttribute("textRotation", TextRotation);
        writer.WriteAttribute("wrapText", WrapText);
        writer.WriteAttribute("indent", Indent);
        writer.WriteAttribute("shrinkToFit", ShrinkToFit);
        writer.AppendLiteral("/>");
    }
}

public enum CellHorizontalAlignment
{
    General,
    Left,
    Center,
    Right,
    Fill,
    Justify,
    CenterContinuous,
    Distributed,
}

public enum CellVerticalAlignment
{
    Top,
    Center,
    Bottom,
    Justify,
    Distributed,
}