using System.Buffers;
using Utf8StringInterpolation;

namespace FastExcelSlim.OpenXml;

public class StyleTable
{
    internal DataFormats DataFormats { get; } = new();

    internal Fonts Fonts { get; } = new();

    internal Fills Fills { get; } = new();

    internal Borders Borders { get; } = new();

    internal CellStyleSheet StyleSheet { get; } = new();

    internal CellStyles CellStyles { get; } = new();

    internal StyleTable()
    {
        Initialize();
    }

    private void Initialize()
    {
        CreateDefaultFont();
        CreateDefaultFills();
        CreateDefaultBorder();
        CreateDefaultCellStyle();
    }

    private void CreateDefaultFont()
    {
        var defaultFont = Fonts.CreateFont();
        defaultFont.SetColor(Font.DefaultFontColor);
        defaultFont.Family = FontFamily.Swiss;
        defaultFont.Scheme = FontScheme.Minor;
    }

    private void CreateDefaultFills()
    {
        Fills.CreateFill().AddPatternFill().PatternType = PatternType.None;
        Fills.CreateFill().AddPatternFill().PatternType = PatternType.DarkGray;
    }

    private void CreateDefaultBorder()
    {
        var border = Borders.CreateBorder();
        border.AddLeft();
        border.AddRight();
        border.AddTop();
        border.AddBottom();
        border.AddDiagonal();
    }

    private void CreateDefaultCellStyle()
    {
        StyleSheet.Initialize();
        CellStyles.CreateCellStyle();
    }

    public DataFormat GetDataFormat(string format)
    {
        return DataFormats.GetDateFormat(format);
    }

    public Font CreateFont()
    {
        return Fonts.CreateFont();
    }

    public Fill CreateFill()
    {
        return Fills.CreateFill();
    }

    public Border CreateBorder()
    {
        return Borders.CreateBorder();
    }

    public CellStyle CreateCellStyle()
    {
        return CellStyles.CreateCellStyle();
    }

    public void Write<TBufferWriter>(scoped ref Utf8StringWriter<TBufferWriter> writer)
        where TBufferWriter : IBufferWriter<byte>
    {
        writer.AppendLiteral(ExcelXml.StylesStart);
        DataFormats.Write(ref writer);
        Fonts.Write(ref writer);
        Fills.Write(ref writer);
        Borders.Write(ref writer);
        StyleSheet.Write(ref writer);
        CellStyles.Write(ref writer);
        writer.AppendLiteral(ExcelXml.StylesEnd);
    }
}