using System.Drawing;
using FastExcelSlim.Extensions;
using Utf8StringInterpolation;

namespace FastExcelSlim.OpenXml;

public class Fill : IdXmlElement
{
    internal Fill(int id) : base(id)
    {
    }

    public PatternFill? PatternFill { get; private set; }

    public PatternFill AddPatternFill()
    {
        PatternFill = new PatternFill();
        return PatternFill;
    }

    internal override void Write<TBufferWriter>(scoped ref Utf8StringWriter<TBufferWriter> writer)
    {
        writer.AppendLiteral("<fill>");
        PatternFill?.Write(ref writer);
        writer.AppendLiteral("</fill>");
    }
}

public class Fills : XmlElementCollection<Fill>
{
    internal Fills() { }
    internal override string Name => "fills";

    public Fill CreateFill()
    {
        var fill = new Fill(Count);
        Add(fill);
        return fill;
    }
}

public class PatternFill : XmlElement
{
    public PatternType? PatternType { get; set; }

    public OpenXmlColor? ForegroundColor { get; private set; }

    public OpenXmlColor? BackgroundColor { get; private set; }

    public void SetForegroundColor(int index)
    {
        ForegroundColor = new OpenXmlColor(index);
    }

    public void SetForegroundColor(Color color)
    {
        ForegroundColor = new OpenXmlColor(color);
    }

    public void SetBackgroundColor(int index)
    {
        BackgroundColor = new OpenXmlColor(index);
    }

    public void SetBackgroundColor(Color color)
    {
        BackgroundColor = new OpenXmlColor(color);
    }

    internal override void Write<TBufferWriter>(scoped ref Utf8StringWriter<TBufferWriter> writer)
    {
        writer.AppendLiteral("<patternFill");
        if (PatternType.HasValue)
        {
            writer.WriteAttribute("patternType", PatternType.Value);
        }

        if (ForegroundColor == null && BackgroundColor == null)
        {
            writer.AppendLiteral("/>");
        }
        else
        {
            writer.AppendLiteral(">");
            ForegroundColor?.Write(ref writer);
            BackgroundColor?.Write(ref writer);
            writer.AppendLiteral("</patternFill>");
        }
    }
}

public enum PatternType
{
    None,
    Solid,
    MediumGray,
    DarkGray,
    LightGray,
    DarkHorizontal,
    DarkVertical,
    DarkDown,
    DarkUp,
    DarkGrid,
    DarkTrellis,
    LightHorizontal,
    LightVertical,
    LightDown,
    LightUp,
    LightGrid,
    LightTrellis,
    Gray125,
    Gray0625,
}
