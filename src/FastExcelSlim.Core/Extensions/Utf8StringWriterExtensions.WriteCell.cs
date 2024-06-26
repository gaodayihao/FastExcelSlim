using System.Buffers;
using Utf8StringInterpolation;

namespace FastExcelSlim.Extensions;

static partial class Utf8StringWriterExtensions
{
    public static void WriteCell<TBufferWriter, T>(
        this scoped ref Utf8StringWriter<TBufferWriter> writer,
        OpenXmlStyles styles,
        string? value, int rowIndex, int columnIndex, string propertyName, scoped ref T entity)
        where T : IOpenXmlWritable<T>
        where TBufferWriter : IBufferWriter<byte>
    {
        var preserveSpace = !string.IsNullOrEmpty(value) && (value.StartsWith(' ') || value.EndsWith(' '));
        var styleIndex = styles.GetCellStyleIndex(propertyName, ref entity);
        writer.AppendLiteral("<c r=\"");
        writer.ConvertXYToCellReference(columnIndex, rowIndex);
        writer.AppendFormat($"\" t=\"str\"{(preserveSpace ? " xml:space=\"preserve\"" : string.Empty)}");
        writer.WriteAttribute("s", styleIndex);
        writer.AppendLiteral("><v>");
        writer.EncodeXml(value);
        writer.AppendLiteral("</v></c>");
    }

    public static void WriteCell<TBufferWriter, T>(
        this scoped ref Utf8StringWriter<TBufferWriter> writer,
        OpenXmlStyles styles,
        bool? value, int rowIndex, int columnIndex, string propertyName, scoped ref T entity)
        where T : IOpenXmlWritable<T>
        where TBufferWriter : IBufferWriter<byte>
    {
        var cellValue = value switch
        {
            true => "1",
            false => "0",
            _ => string.Empty
        };
        var styleIndex = styles.GetCellStyleIndex(propertyName, ref entity);
        writer.AppendLiteral("<c r=\"");
        writer.ConvertXYToCellReference(columnIndex, rowIndex);
        writer.AppendLiteral("\" t=\"b\"");
        writer.WriteAttribute("s", styleIndex);
        writer.AppendFormat($"><v>{cellValue}</v></c>");
    }

    public static void WriteCell<TBufferWriter, T>(
        this scoped ref Utf8StringWriter<TBufferWriter> writer,
        OpenXmlStyles styles,
        Guid? value, int rowIndex, int columnIndex, string propertyName, scoped ref T entity)
        where T : IOpenXmlWritable<T>
        where TBufferWriter : IBufferWriter<byte>
    {
        var styleIndex = styles.GetCellStyleIndex(propertyName, ref entity);
        writer.AppendLiteral("<c r=\"");
        writer.ConvertXYToCellReference(columnIndex, rowIndex);
        writer.AppendLiteral("\" t=\"str\"");
        writer.WriteAttribute("s", styleIndex);
        writer.AppendFormat($"><v>{value}</v></c>");
    }

    private static double CorrectDateTimeValue(DateTime value)
    {
        var oaDate = value.ToOADate();

        // Excel says 1900 was a leap year  :( Replicate an incorrect behavior thanks
        // to Lotus 1-2-3 decision from 1983...
        // https://github.com/ClosedXML/ClosedXML/blob/develop/ClosedXML/Extensions/DateTimeExtensions.cs#L45
        const int nonExistent1900Feb29SerialDate = 60;
        if (oaDate <= nonExistent1900Feb29SerialDate)
        {
            oaDate -= 1;
        }

        return oaDate;
    }

    public static void WriteCell<TBufferWriter, T>(
        this scoped ref Utf8StringWriter<TBufferWriter> writer,
        OpenXmlStyles styles,
        DateOnly? value, int rowIndex, int columnIndex, string propertyName, scoped ref T entity)
        where T : IOpenXmlWritable<T>
        where TBufferWriter : IBufferWriter<byte>
    {
        var styleIndex = styles.GetDateTimeCellStyleIndex(propertyName, ref entity);
        writer.AppendLiteral("<c r=\"");
        writer.ConvertXYToCellReference(columnIndex, rowIndex);
        writer.AppendLiteral("\"");
        writer.WriteAttribute("s", styleIndex);
        writer.AppendLiteral("><v>");
        if (value.HasValue)
        {
            var oaDate = CorrectDateTimeValue(value.Value.ToDateTime(TimeOnly.MinValue));
            writer.AppendFormatted(oaDate);
        }
        writer.AppendLiteral("</v></c>");
    }
    
    public static void WriteCell<TBufferWriter, T>(
        this scoped ref Utf8StringWriter<TBufferWriter> writer,
        OpenXmlStyles styles,
        TimeOnly? value, int rowIndex, int columnIndex, string propertyName, scoped ref T entity)
        where T : IOpenXmlWritable<T>
        where TBufferWriter : IBufferWriter<byte>
    {
        var styleIndex = styles.GetCellStyleIndex(propertyName, ref entity);
        writer.AppendLiteral("<c r=\"");
        writer.ConvertXYToCellReference(columnIndex, rowIndex);
        writer.AppendLiteral("\" t=\"str\"");
        writer.WriteAttribute("s", styleIndex);
        writer.AppendFormat($"><v>{value:HH:mm:ss}</v></c>");
    }

    public static void WriteCell<TBufferWriter, T>(
        this scoped ref Utf8StringWriter<TBufferWriter> writer,
        OpenXmlStyles styles,
        DateTime? value, int rowIndex, int columnIndex, string propertyName, scoped ref T entity)
        where T : IOpenXmlWritable<T>
        where TBufferWriter : IBufferWriter<byte>
    {
        var styleIndex = styles.GetDateTimeCellStyleIndex(propertyName, ref entity);
        writer.AppendLiteral("<c r=\"");
        writer.ConvertXYToCellReference(columnIndex, rowIndex);
        writer.AppendLiteral("\"");
        writer.WriteAttribute("s", styleIndex);
        writer.AppendLiteral("><v>");
        if (value.HasValue)
        {
            var oaDate = CorrectDateTimeValue(value.Value);
            writer.AppendFormatted(oaDate);
        }
        writer.AppendLiteral("</v></c>");
    }

    public static void WriteCell<TBufferWriter, T>(
        this scoped ref Utf8StringWriter<TBufferWriter> writer,
        OpenXmlStyles styles,
        byte? value, int rowIndex, int columnIndex, string propertyName, scoped ref T entity)
        where T : IOpenXmlWritable<T>
        where TBufferWriter : IBufferWriter<byte>
    {
        var styleIndex = styles.GetCellStyleIndex(propertyName, ref entity);
        writer.AppendLiteral("<c r=\"");
        writer.ConvertXYToCellReference(columnIndex, rowIndex);
        writer.AppendLiteral("\"");
        writer.WriteAttribute("s", styleIndex);
        writer.AppendFormat($"><v>{value}</v></c>");
    }

    public static void WriteCell<TBufferWriter, T>(
        this scoped ref Utf8StringWriter<TBufferWriter> writer,
        OpenXmlStyles styles,
        decimal? value, int rowIndex, int columnIndex, string propertyName, scoped ref T entity)
        where T : IOpenXmlWritable<T>
        where TBufferWriter : IBufferWriter<byte>
    {
        var styleIndex = styles.GetCellStyleIndex(propertyName, ref entity);
        writer.AppendLiteral("<c r=\"");
        writer.ConvertXYToCellReference(columnIndex, rowIndex);
        writer.AppendLiteral("\"");
        writer.WriteAttribute("s", styleIndex);
        writer.AppendFormat($"><v>{value}</v></c>");
    }

    public static void WriteCell<TBufferWriter, T>(
        this scoped ref Utf8StringWriter<TBufferWriter> writer,
        OpenXmlStyles styles,
        double? value, int rowIndex, int columnIndex, string propertyName, scoped ref T entity)
        where T : IOpenXmlWritable<T>
        where TBufferWriter : IBufferWriter<byte>
    {
        var styleIndex = styles.GetCellStyleIndex(propertyName, ref entity);
        writer.AppendLiteral("<c r=\"");
        writer.ConvertXYToCellReference(columnIndex, rowIndex);
        writer.AppendLiteral("\"");
        writer.WriteAttribute("s", styleIndex);
        writer.AppendFormat($"><v>{value}</v></c>");
    }

    public static void WriteCell<TBufferWriter, T>(
        this scoped ref Utf8StringWriter<TBufferWriter> writer,
        OpenXmlStyles styles,
        short? value, int rowIndex, int columnIndex, string propertyName, scoped ref T entity)
        where T : IOpenXmlWritable<T>
        where TBufferWriter : IBufferWriter<byte>
    {
        var styleIndex = styles.GetCellStyleIndex(propertyName, ref entity);
        writer.AppendLiteral("<c r=\"");
        writer.ConvertXYToCellReference(columnIndex, rowIndex);
        writer.AppendLiteral("\"");
        writer.WriteAttribute("s", styleIndex);
        writer.AppendFormat($"><v>{value}</v></c>");
    }

    public static void WriteCell<TBufferWriter, T>(
        this scoped ref Utf8StringWriter<TBufferWriter> writer,
        OpenXmlStyles styles,
        int? value, int rowIndex, int columnIndex, string propertyName, scoped ref T entity)
        where T : IOpenXmlWritable<T>
        where TBufferWriter : IBufferWriter<byte>
    {
        var styleIndex = styles.GetCellStyleIndex(propertyName, ref entity);
        writer.AppendLiteral("<c r=\"");
        writer.ConvertXYToCellReference(columnIndex, rowIndex);
        writer.AppendLiteral("\"");
        writer.WriteAttribute("s", styleIndex);
        writer.AppendFormat($"><v>{value}</v></c>");
    }

    public static void WriteCell<TBufferWriter, T>(
        this scoped ref Utf8StringWriter<TBufferWriter> writer,
        OpenXmlStyles styles,
        long? value, int rowIndex, int columnIndex, string propertyName, scoped ref T entity)
        where T : IOpenXmlWritable<T>
        where TBufferWriter : IBufferWriter<byte>
    {
        var styleIndex = styles.GetCellStyleIndex(propertyName, ref entity);
        writer.AppendLiteral("<c r=\"");
        writer.ConvertXYToCellReference(columnIndex, rowIndex);
        writer.AppendLiteral("\"");
        writer.WriteAttribute("s", styleIndex);
        writer.AppendFormat($"><v>{value}</v></c>");
    }

    public static void WriteCell<TBufferWriter, T>(
        this scoped ref Utf8StringWriter<TBufferWriter> writer,
        OpenXmlStyles styles,
        sbyte? value, int rowIndex, int columnIndex, string propertyName, scoped ref T entity)
        where T : IOpenXmlWritable<T>
        where TBufferWriter : IBufferWriter<byte>
    {
        var styleIndex = styles.GetCellStyleIndex(propertyName, ref entity);
        writer.AppendLiteral("<c r=\"");
        writer.ConvertXYToCellReference(columnIndex, rowIndex);
        writer.AppendLiteral("\"");
        writer.WriteAttribute("s", styleIndex);
        writer.AppendFormat($"><v>{value}</v></c>");
    }

    public static void WriteCell<TBufferWriter, T>(
        this scoped ref Utf8StringWriter<TBufferWriter> writer,
        OpenXmlStyles styles,
        float? value, int rowIndex, int columnIndex, string propertyName, scoped ref T entity)
        where T : IOpenXmlWritable<T>
        where TBufferWriter : IBufferWriter<byte>
    {
        var styleIndex = styles.GetCellStyleIndex(propertyName, ref entity);
        writer.AppendLiteral("<c r=\"");
        writer.ConvertXYToCellReference(columnIndex, rowIndex);
        writer.AppendLiteral("\"");
        writer.WriteAttribute("s", styleIndex);
        writer.AppendFormat($"><v>{value}</v></c>");
    }

    public static void WriteCell<TBufferWriter, T>(
        this scoped ref Utf8StringWriter<TBufferWriter> writer,
        OpenXmlStyles styles,
        ushort? value, int rowIndex, int columnIndex, string propertyName, scoped ref T entity)
        where T : IOpenXmlWritable<T>
        where TBufferWriter : IBufferWriter<byte>
    {
        var styleIndex = styles.GetCellStyleIndex(propertyName, ref entity);
        writer.AppendLiteral("<c r=\"");
        writer.ConvertXYToCellReference(columnIndex, rowIndex);
        writer.AppendLiteral("\"");
        writer.WriteAttribute("s", styleIndex);
        writer.AppendFormat($"><v>{value}</v></c>");
    }

    public static void WriteCell<TBufferWriter, T>(
        this scoped ref Utf8StringWriter<TBufferWriter> writer,
        OpenXmlStyles styles,
        uint? value, int rowIndex, int columnIndex, string propertyName, scoped ref T entity)
        where T : IOpenXmlWritable<T>
        where TBufferWriter : IBufferWriter<byte>
    {
        var styleIndex = styles.GetCellStyleIndex(propertyName, ref entity);
        writer.AppendLiteral("<c r=\"");
        writer.ConvertXYToCellReference(columnIndex, rowIndex);
        writer.AppendLiteral("\"");
        writer.WriteAttribute("s", styleIndex);
        writer.AppendFormat($"><v>{value}</v></c>");
    }

    public static void WriteCell<TBufferWriter, T>(
        this scoped ref Utf8StringWriter<TBufferWriter> writer,
        OpenXmlStyles styles,
        ulong? value, int rowIndex, int columnIndex, string propertyName, scoped ref T entity)
        where T : IOpenXmlWritable<T>
        where TBufferWriter : IBufferWriter<byte>
    {
        var styleIndex = styles.GetCellStyleIndex(propertyName, ref entity);
        writer.AppendLiteral("<c r=\"");
        writer.ConvertXYToCellReference(columnIndex, rowIndex);
        writer.AppendLiteral("\"");
        writer.WriteAttribute("s", styleIndex);
        writer.AppendFormat($"><v>{value}</v></c>");
    }

}