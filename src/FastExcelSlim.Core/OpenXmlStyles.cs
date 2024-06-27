using System.Buffers;
using System.Drawing;
using FastExcelSlim.OpenXml;
using Utf8StringInterpolation;

namespace FastExcelSlim;

public abstract class OpenXmlStyles
{
    protected StyleTable StyleTable { get; } = new();

    protected const int NoneStyleIndex = 0;

    public abstract int GetCellStyleIndex<T>(string propertyName, scoped ref T entity);

    public abstract int GetHeaderStyleIndex(string propertyName);

    public abstract int GetDateTimeCellStyleIndex<T>(string propertyName, scoped ref T entity);

    internal void Write<TBufferWriter>(scoped ref Utf8StringWriter<TBufferWriter> writer)
        where TBufferWriter : IBufferWriter<byte>
    {
        StyleTable.Write(ref writer);
    }
}

public class NoneStyles : OpenXmlStyles
{
    private readonly int _defaultDateTimeCellStyleIndex;

    internal NoneStyles()
    {
        var defaultDateTimeCellStyle = StyleTable.CreateCellStyle();
        var format = StyleTable.GetDataFormat(BuiltinFormats.DefaultDateTimeFormat);
        defaultDateTimeCellStyle.DataFormat = format;
        defaultDateTimeCellStyle.AddAlignment().Vertical = CellVerticalAlignment.Center;

        _defaultDateTimeCellStyleIndex = defaultDateTimeCellStyle.Id;
    }

    public override int GetCellStyleIndex<T>(string propertyName, scoped ref T entity)
    {
        return NoneStyleIndex;
    }

    public override int GetHeaderStyleIndex(string propertyName)
    {
        return NoneStyleIndex;
    }

    public override int GetDateTimeCellStyleIndex<T>(string propertyName, scoped ref T entity)
    {
        return _defaultDateTimeCellStyleIndex;
    }
}

public class DefaultStyles : OpenXmlStyles
{
    public static readonly DefaultStyles Default = new();
    public static readonly OpenXmlStyles None = new NoneStyles();

    protected Border DefaultBorder { get; }

    protected CellStyle DefaultCellStyle { get; }
    protected CellStyle DefaultHeaderStyle { get; }
    protected CellStyle DefaultDateTimeCellStyle { get; }

    private DefaultStyles()
    {
        DefaultBorder = StyleTable.CreateBorder();
        DefaultBorder.AddTop().Style = BorderStyle.Thin;
        DefaultBorder.AddRight().Style = BorderStyle.Thin;
        DefaultBorder.AddLeft().Style = BorderStyle.Thin;
        DefaultBorder.AddBottom().Style = BorderStyle.Thin;
        DefaultBorder.Top!.SetColor(IndexedColors.Automatic);
        DefaultBorder.Right!.SetColor(IndexedColors.Automatic);
        DefaultBorder.Left!.SetColor(IndexedColors.Automatic);
        DefaultBorder.Bottom!.SetColor(IndexedColors.Automatic);

        // ReSharper disable VirtualMemberCallInConstructor
        DefaultCellStyle = CreateDefaultCellStyle();
        DefaultHeaderStyle = CreateDefaultHeaderStyle();
        DefaultDateTimeCellStyle = CreateDefaultDateTimeStyle();
    }

    public override int GetCellStyleIndex<T>(string propertyName, scoped ref T entity)
    {
        return DefaultCellStyle.Id;
    }

    public override int GetHeaderStyleIndex(string propertyName)
    {
        return DefaultHeaderStyle.Id;
    }

    public override int GetDateTimeCellStyleIndex<T>(string propertyName, scoped ref T entity)
    {
        return DefaultDateTimeCellStyle.Id;
    }

    protected virtual CellStyle CreateDefaultDateTimeStyle()
    {
        var defaultDateTimeCellStyle = StyleTable.CreateCellStyle();
        defaultDateTimeCellStyle.Border = DefaultBorder;
        var format = StyleTable.GetDataFormat(BuiltinFormats.DefaultDateTimeFormat);
        defaultDateTimeCellStyle.DataFormat = format;
        defaultDateTimeCellStyle.AddAlignment().Vertical = CellVerticalAlignment.Center;
        return defaultDateTimeCellStyle;
    }

    protected virtual CellStyle CreateDefaultHeaderStyle()
    {
        var headerCellStyle = StyleTable.CreateCellStyle();
        headerCellStyle.Border = DefaultBorder;
        headerCellStyle.Fill = StyleTable.CreateFill();
        var patternFill = headerCellStyle.Fill.AddPatternFill();
        patternFill.PatternType = PatternType.Solid;
        patternFill.SetForegroundColor(Color.FromArgb(77, 113, 195));
        patternFill.SetBackgroundColor(IndexedColors.Automatic);
        var headerFont = StyleTable.CreateFont();
        headerFont.IsBold = true;
        headerFont.SetColor(IndexedColors.White);
        headerCellStyle.Font = headerFont;
        return headerCellStyle;
    }

    protected virtual CellStyle CreateDefaultCellStyle()
    {
        var cellStyle = StyleTable.CreateCellStyle();
        cellStyle.Border = DefaultBorder;
        return cellStyle;
    }
}