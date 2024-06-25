using System.Drawing;
using FastExcelSlim.Extensions;
using Utf8StringInterpolation;

namespace FastExcelSlim.OpenXml;

public class Border : IdXmlElement
{
    public bool DiagonalUp { get; set; }

    public bool DiagonalDown { get; set; }

    public bool Outline { get; set; }

    public BorderPr? Left { get; private set; }

    public BorderPr? Right { get; private set; }

    public BorderPr? Top { get; private set; }

    public BorderPr? Bottom { get; private set; }

    public BorderPr? Diagonal { get; private set; }

    public BorderPr? Vertical { get; private set; }

    public BorderPr? Horizontal { get; private set; }

    public BorderPr AddLeft()
    {
        return Left ??= new BorderPr("left");
    }

    public BorderPr AddRight()
    {
        return Right ??= new BorderPr("right");
    }

    public BorderPr AddTop()
    {
        return Top ??= new BorderPr("top");
    }

    public BorderPr AddBottom()
    {
        return Bottom ??= new BorderPr("bottom");
    }

    public BorderPr AddDiagonal()
    {
        return Diagonal ??= new BorderPr("diagonal");
    }

    public BorderPr AddVertical()
    {
        return Vertical ??= new BorderPr("vertical");
    }

    public BorderPr AddHorizontal()
    {
        return Horizontal ??= new BorderPr("horizontal");
    }

    internal Border(int id) : base(id)
    {
    }

    internal override void Write<TBufferWriter>(scoped ref Utf8StringWriter<TBufferWriter> writer)
    {
        writer.AppendLiteral("<border");
        writer.WriteAttribute("diagonalUp", DiagonalUp);
        writer.WriteAttribute("diagonalDown", DiagonalDown);
        writer.WriteAttribute("outline", Outline);
        writer.AppendLiteral(">");
        Left?.Write(ref writer);
        Right?.Write(ref writer);
        Top?.Write(ref writer);
        Bottom?.Write(ref writer);
        Diagonal?.Write(ref writer);
        Vertical?.Write(ref writer);
        Horizontal?.Write(ref writer);
        writer.AppendLiteral("</border>");
    }
}

public class Borders : XmlElementCollection<Border>
{
    internal Borders() { }

    internal override string Name => "borders";

    public Border CreateBorder()
    {
        var border = new Border(Count);
        Add(border);
        return border;
    }
}

public class BorderPr(string nodeName) : XmlElement
{
    public BorderStyle Style { get; set; } = BorderStyle.None;

    public OpenXmlColor? Color { get; private set; }

    public void SetColor(int index)
    {
        Color = new OpenXmlColor(index);
    }

    public void SetColor(Color color)
    {
        Color = new OpenXmlColor(color);
    }

    internal override void Write<TBufferWriter>(scoped ref Utf8StringWriter<TBufferWriter> writer)
    {
        writer.AppendFormat($"<{nodeName}");
        if (Style != BorderStyle.None)
        {
            writer.WriteAttribute("style", Style);
        }
        if (Color != null)
        {
            writer.AppendLiteral(">");
            Color.Write(ref writer);
            writer.AppendFormat($"</{nodeName}>");
        }
        else
        {
            writer.AppendLiteral("/>");
        }
    }
}

public enum BorderStyle
{
    None,
    Thin,
    Medium,
    Dashed,
    Dotted,
    Thick,
    Double,
    Hair,
    MediumDashed,
    DashDot,
    MediumDashDot,
    DashDotDot,
    MediumDashDotDot,
    SlantDashDot,
}