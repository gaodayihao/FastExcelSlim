﻿using System.Buffers;
using Utf8StringInterpolation;

namespace FastExcelSlim.OpenXml;

internal class OpenXmlWorkbook
{
    private readonly List<OpenXmlSheet> _sheets = new(1);

    public OpenXmlSheet<T> CreateSheet<T>(IEnumerable<T> values, OpenXmlStyles<T> styles)
        where T : IOpenXmlWritable<T>
    {
        var sheet = new OpenXmlSheet<T>(_sheets.Count + 1, values, styles);
        _sheets.Add(sheet);
        return sheet;
    }

    public void WriteRelationships(scoped ref ZipEntryWriterWrapper wrapper)
    {
        wrapper.Writer.AppendLiteral(ExcelXml.WorkbookRelationshipStart);

        int lastId = 0;
        foreach (var sheet in _sheets)
        {
            wrapper.Writer.AppendFormat($"<Relationship Id=\"rId{sheet.Id}\" Type=\"{ExcelNamespace.Worksheet}\" Target=\"worksheets/sheet{sheet.Id}.xml\" />");
            lastId = sheet.Id;
            wrapper.CheckAndFlush();
        }
        wrapper.Writer.AppendFormat($"<Relationship Id=\"rId{++lastId}\" Type=\"{ExcelNamespace.SharedStrings}\" Target=\"sharedStrings.xml\" />");
        wrapper.Writer.AppendFormat($"<Relationship Id=\"rId{++lastId}\" Type=\"{ExcelNamespace.Styles}\" Target=\"styles.xml\" />");
        wrapper.Writer.AppendLiteral(ExcelXml.WorkbookRelationshipEnd);
    }

    public void Write(scoped ref ZipEntryWriterWrapper wrapper)
    {
        wrapper.Writer.AppendLiteral(ExcelXml.WorkbookStart);
        wrapper.Writer.AppendLiteral(ExcelXml.WorkbookPr);
        wrapper.Writer.AppendLiteral(ExcelXml.WorkbookViews);
        WriteSheets(ref wrapper);
        wrapper.Writer.AppendLiteral(ExcelXml.WorkbookEnd);
    }

    private void WriteSheets(scoped ref ZipEntryWriterWrapper wrapper)
    {
        wrapper.Writer.AppendLiteral("<sheets>");
        foreach (var sheet in _sheets)
        {
            wrapper.Writer.AppendFormat($"<sheet name=\"{sheet.Name}\" sheetId=\"{sheet.Id}\" r:id=\"rId{sheet.Id}\"></sheet>");
            wrapper.CheckAndFlush();
        }
        wrapper.Writer.AppendLiteral("</sheets>");
    }
}