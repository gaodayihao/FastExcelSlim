using System.Drawing;
using FastExcelSlim.OpenXml;

namespace FastExcelSlim;

internal class SampleStyles : OpenXmlStyles
{
    private readonly CellStyle _headerStyle;
    private readonly CellStyle _singularLineStyle;
    private readonly CellStyle _pluralLineStyle;
    private readonly CellStyle _singularLineDateTimeStyle;
    private readonly CellStyle _pluralLineDateTimeStyle;

    public SampleStyles()
    {
        _headerStyle = StyleTable.CreateCellStyle();
        var headerFont = StyleTable.CreateFont();
        headerFont.IsBold = true;
        _headerStyle.Font = headerFont;

        var singularLineCellFill = StyleTable.CreateFill();
        var fill = singularLineCellFill.AddPatternFill();
        fill.PatternType = PatternType.Solid;
        fill.SetForegroundColor(Color.Gainsboro);

        var pluralLineFill = StyleTable.CreateFill();
        fill = pluralLineFill.AddPatternFill();
        fill.PatternType = PatternType.Solid;
        fill.SetForegroundColor(Color.CornflowerBlue);

        var dateTimeFormat = StyleTable.GetDataFormat("yyyy-mm-dd");

        _singularLineDateTimeStyle = StyleTable.CreateCellStyle();
        _singularLineDateTimeStyle.DataFormat = dateTimeFormat;
        _singularLineDateTimeStyle.AddAlignment().Vertical = CellVerticalAlignment.Center;
        _singularLineDateTimeStyle.Fill = singularLineCellFill;

        _pluralLineDateTimeStyle = StyleTable.CreateCellStyle();
        _pluralLineDateTimeStyle.DataFormat = dateTimeFormat;
        _pluralLineDateTimeStyle.AddAlignment().Vertical = CellVerticalAlignment.Center;
        _pluralLineDateTimeStyle.Fill = pluralLineFill;

        _singularLineStyle = StyleTable.CreateCellStyle();
        _singularLineStyle.Fill = singularLineCellFill;

        _pluralLineStyle = StyleTable.CreateCellStyle();
        _pluralLineStyle.Fill = pluralLineFill;
    }

    public override int GetCellStyleIndex<T>(string propertyName, int rowIndex, ref T entity)
    {
        if (rowIndex % 2 == 0)
        {
            return _pluralLineStyle.Id;
        }

        return _singularLineStyle.Id;
    }

    public override int GetHeaderStyleIndex(string propertyName)
    {
        return _headerStyle.Id;
    }

    public override int GetDateTimeCellStyleIndex<T>(string propertyName, int rowIndex, ref T entity)
    {
        if (rowIndex % 2 == 0)
        {
            return _pluralLineDateTimeStyle.Id;
        }

        return _singularLineDateTimeStyle.Id;
    }
}