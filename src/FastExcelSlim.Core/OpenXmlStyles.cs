using System.Buffers;
using FastExcelSlim.OpenXml;
using Utf8StringInterpolation;

namespace FastExcelSlim;

public abstract class OpenXmlStyles
{
    protected StyleTable StyleTable { get; } = new();

    internal void Write<TBufferWriter>(scoped ref Utf8StringWriter<TBufferWriter> writer)
        where TBufferWriter : IBufferWriter<byte>
    {
        StyleTable.Write(ref writer);
    }
}

public abstract class OpenXmlStyles<T> : OpenXmlStyles
{
    protected const int DefaultCellStyleIndex = 0;

    public abstract int GetCellStyleIndex(string propertyName, scoped ref T entity);

    public abstract int GetHeaderStyleIndex(string propertyName);

    public abstract int GetDateTimeCellStyleIndex(string propertyName, scoped ref T entity);
}

public class DefaultStyles<T> : OpenXmlStyles<T>
{
    public CellStyle DefaultHeaderStyle { get; set; }
    protected CellStyle DefaultDateTimeCellStyle { get; set; }

    public DefaultStyles()
    {
        // ReSharper disable VirtualMemberCallInConstructor
        DefaultDateTimeCellStyle = CreateDefaultDateTimeStyle();
        DefaultHeaderStyle = CreateDefaultHeaderStyle();
    }

    public override int GetCellStyleIndex(string propertyName, scoped ref T entity)
    {
        return DefaultCellStyleIndex;
    }

    public override int GetHeaderStyleIndex(string propertyName)
    {
        return DefaultHeaderStyle.Id;
    }

    public override int GetDateTimeCellStyleIndex(string propertyName, scoped ref T entity)
    {
        return DefaultDateTimeCellStyle.Id;
    }

    protected virtual CellStyle CreateDefaultDateTimeStyle()
    {
        var defaultDateTimeCellStyle = StyleTable.CreateCellStyle();
        var format = StyleTable.GetDataFormat(BuiltinFormats.DefaultDateTimeFormat);
        defaultDateTimeCellStyle.DataFormat = format;
        defaultDateTimeCellStyle.AddAlignment().Vertical = CellVerticalAlignment.Center;
        return defaultDateTimeCellStyle;
    }

    protected virtual CellStyle CreateDefaultHeaderStyle()
    {
        var headerCellStyle = StyleTable.CreateCellStyle();
        var headerFont = StyleTable.CreateFont();
        headerFont.IsBold = true;
        headerCellStyle.Font = headerFont;
        return headerCellStyle;
    }
}