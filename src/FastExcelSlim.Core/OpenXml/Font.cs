using System.Buffers;
using System.Drawing;
using FastExcelSlim.Extensions;
using Utf8StringInterpolation;

namespace FastExcelSlim.OpenXml;

public class Font : IdXmlElement
{
    internal Font(int id) : base(id) { }

    public const int DefaultFontColor = IndexedColors.Black;
    public const int DefaultFontSize = 11;
    public const string DefaultFontName = "Calibri";

    public int Size { get; set; } = DefaultFontSize;

    public string Name { get; set; } = DefaultFontName;

    public bool IsBold { get; set; }

    public bool IsItalics { get; set; }

    public bool IsStrikeout { get; set; }

    public OpenXmlColor? Color { get; private set; }

    public FontFamily? Family { get; set; }

    public FontScheme? Scheme { get; set; }

    public UnderlineValues? Underline { get; set; }

    public VerticalAlignRun? VerticalAlign { get; set; }

    public FontCharset? FontCharset { get; set; }

    public void SetColor(int colorIndex)
    {
        Color = new OpenXmlColor(colorIndex);
    }

    public void SetColor(Color color)
    {
        Color = new OpenXmlColor(color);
    }

    private void WriteSize<TBufferWriter>(scoped ref Utf8StringWriter<TBufferWriter> writer)
        where TBufferWriter : IBufferWriter<byte>
    {
        writer.AppendLiteral("<sz");
        writer.WriteAttribute("val", Size, true);
        writer.AppendLiteral("/>");
    }

    private void WriteName<TBufferWriter>(scoped ref Utf8StringWriter<TBufferWriter> writer)
        where TBufferWriter : IBufferWriter<byte>
    {
        if (!string.IsNullOrEmpty(Name))
        {
            writer.AppendLiteral("<name");
            writer.WriteAttribute("val", Name);
            writer.AppendLiteral("/>");
        }
    }

    private void WriteFamily<TBufferWriter>(scoped ref Utf8StringWriter<TBufferWriter> writer)
        where TBufferWriter : IBufferWriter<byte>
    {
        if (Family.HasValue)
        {
            writer.AppendLiteral("<family");
            writer.WriteAttribute("val", Family.Value, "D");
            writer.AppendLiteral("/>");
        }
    }

    private void WriteFontCharset<TBufferWriter>(scoped ref Utf8StringWriter<TBufferWriter> writer)
        where TBufferWriter : IBufferWriter<byte>
    {
        if (FontCharset.HasValue)
        {
            writer.AppendLiteral("<charset");
            writer.WriteAttribute("val", FontCharset.Value, "D");
            writer.AppendLiteral("/>");
        }
    }

    private void WriteScheme<TBufferWriter>(scoped ref Utf8StringWriter<TBufferWriter> writer)
        where TBufferWriter : IBufferWriter<byte>
    {
        if (Scheme.HasValue)
        {
            writer.AppendLiteral("<scheme");
            writer.WriteAttribute("val", Scheme.Value);
            writer.AppendLiteral("/>");
        }
    }

    private void WriteUnderline<TBufferWriter>(scoped ref Utf8StringWriter<TBufferWriter> writer)
        where TBufferWriter : IBufferWriter<byte>
    {
        if (Underline.HasValue)
        {
            writer.AppendLiteral("<u");
            writer.WriteAttribute("val", Underline.Value);
            writer.AppendLiteral("/>");
        }
    }

    private void WriteVerticalAlignRun<TBufferWriter>(scoped ref Utf8StringWriter<TBufferWriter> writer)
        where TBufferWriter : IBufferWriter<byte>
    {
        if (VerticalAlign.HasValue)
        {
            writer.AppendLiteral("<vertAlign");
            writer.WriteAttribute("val", VerticalAlign.Value);
            writer.AppendLiteral("/>");
        }
    }

    internal override void Write<TBufferWriter>(scoped ref Utf8StringWriter<TBufferWriter> writer)
    {
        writer.AppendLiteral("<font>");
        if (IsBold) writer.AppendLiteral("<b />");
        if (IsItalics) writer.AppendLiteral("<i />");
        if (IsStrikeout) writer.AppendLiteral("<strike />");
        WriteSize(ref writer);
        Color?.Write(ref writer);
        WriteName(ref writer);
        WriteFamily(ref writer);
        WriteScheme(ref writer);
        WriteUnderline(ref writer);
        WriteFontCharset(ref writer);
        WriteVerticalAlignRun(ref writer);
        writer.AppendLiteral("</font>");
    }
}

public enum VerticalAlignRun
{
    Baseline,
    Superscript,
    Subscript,
}

public enum FontFamily
{
    NotApplicable = 0,
    Roman = 1,
    Swiss = 2,
    Modern = 3,
    Script = 4,
    Decorative = 5
}

public enum FontScheme
{
    None = 1,
    Major = 2,
    Minor = 3
}

public enum UnderlineValues
{
    None,
    Single,
    Double,
    SingleAccounting,
    DoubleAccounting
}

public enum FontCharset
{
    Ansi = 0,
    Default = 1,
    Symbol = 2,
    Mac = 77,
    ShiftJis = 128,
    Hangul = 129,
    JoHab = 130,
    Gb2312 = 134,
    ChineseBig5 = 136,
    Greek = 161,
    Turkish = 162,
    Vietnamese = 163,
    Hebrew = 177,
    Arabic = 178,
    Baltic = 186,
    Russian = 204,
    Thai = 222,
    EastEurope = 238,
    Oem = 255,
}

public class Fonts : XmlElementCollection<Font>
{
    internal Fonts() { }
    internal override string Name => "fonts";

    public Font CreateFont()
    {
        var font = new Font(Count);
        Add(font);
        return font;
    }
}
