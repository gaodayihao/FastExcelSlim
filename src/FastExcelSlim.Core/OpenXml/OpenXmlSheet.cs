using System.Buffers;
using Utf8StringInterpolation;

namespace FastExcelSlim.OpenXml;

internal abstract class OpenXmlSheet(int id)
{
    public int Id { get; } = id;

    public string Name { get; protected set; }

    public string Path => $"xl/worksheets/sheet{Id}.xml";
}

internal class OpenXmlSheet<T> : OpenXmlSheet
{
    private readonly IEnumerable<T> _values;
    private readonly OpenXmlStyles<T> _styles;
    private readonly SheetDimension _sheetDimension;
    private readonly IOpenXmlFormatter<T> _formatter;

    public OpenXmlSheet(int id, IEnumerable<T> values, OpenXmlStyles<T> styles) : base(id)
    {
        _values = values;
        _styles = styles;
        _formatter = OpenXmlFormatterProvider.GetFormatter<T>();
        var maxColumn = _formatter.ColumnCount;
        var maxRow = 1;
        if (values is ICollection<T> collection)
        {
            maxRow = collection.Count + 1;
        }

        _sheetDimension = new SheetDimension(maxRow, maxColumn);

        Name = _formatter.SheetName ?? $"sheet{id}";
    }

    public void Write(scoped ref ZipEntryWriterWrapper wrapper)
    {
        wrapper.Writer.AppendLiteral(ExcelXml.SheetStart);
        _sheetDimension.Write(ref wrapper.Writer);
        wrapper.Writer.AppendLiteral(ExcelXml.SheetViews);
        wrapper.Writer.AppendLiteral(ExcelXml.SheetFormatPr);
        wrapper.CheckAndFlush();
        WriteColumns(ref wrapper);
        WriteSheetData(ref wrapper);
        wrapper.Writer.AppendLiteral(ExcelXml.SheetPageMargins);
        wrapper.Writer.AppendLiteral(ExcelXml.SheetEnd);
    }

    private void WriteColumns(scoped ref ZipEntryWriterWrapper wrapper)
    {
        wrapper.Writer.AppendLiteral("<cols>");
        _formatter.WriteColumns(ref wrapper.Writer);
        wrapper.Writer.AppendLiteral("</cols>");
        wrapper.CheckAndFlush();
    }

    private void WriteSheetData(scoped ref ZipEntryWriterWrapper wrapper)
    {
        wrapper.Writer.AppendLiteral("<sheetData>");
        WriteHeaderRow(ref wrapper);

        var rowIndex = 2;
        foreach (var value in _values)
        {
            var cellValue = value;
            wrapper.Writer.AppendFormat($"<row r=\"{rowIndex}\">");
            _formatter.WriteCell(ref wrapper.Writer, _styles, rowIndex++, ref cellValue);
            wrapper.Writer.AppendLiteral("</row>");
            wrapper.CheckAndFlush();
        }
        wrapper.Writer.AppendLiteral("</sheetData>");
    }

    private void WriteHeaderRow(scoped ref ZipEntryWriterWrapper wrapper)
    {
        wrapper.Writer.AppendLiteral("<row r=\"1\">");
        _formatter.WriteHeaders(ref wrapper.Writer, _styles);
        wrapper.Writer.AppendLiteral("</row>");
        wrapper.CheckAndFlush();
    }
}