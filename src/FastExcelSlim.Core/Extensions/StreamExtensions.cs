using FastExcelSlim.OpenXml;

namespace FastExcelSlim;

public static class StreamExtensions
{
    public static void SaveToExcel<T>(this Stream stream, IEnumerable<T> items, OpenXmlStyles? styles = null)
    {
        var writer = new OpenXmlWorkbookWriter(stream, styles);
        writer.CreateSheet(items);
        writer.Save();
    }
    public static void SaveToExcel<T1, T2>(this Stream stream, IEnumerable<T1> items1, IEnumerable<T2> items2, OpenXmlStyles? styles = null)
    {
        var writer = new OpenXmlWorkbookWriter(stream, styles);
        writer.CreateSheet(items1);
        writer.CreateSheet(items2);
        writer.Save();
    }
    public static void SaveToExcel<T1, T2, T3>(this Stream stream, IEnumerable<T1> items1, IEnumerable<T2> items2, IEnumerable<T3> items3, OpenXmlStyles? styles = null)
    {
        var writer = new OpenXmlWorkbookWriter(stream, styles);
        writer.CreateSheet(items1);
        writer.CreateSheet(items2);
        writer.CreateSheet(items3);
        writer.Save();
    }
    public static void SaveToExcel<T1, T2, T3, T4>(this Stream stream, IEnumerable<T1> items1, IEnumerable<T2> items2, IEnumerable<T3> items3, IEnumerable<T4> items4, OpenXmlStyles? styles = null)
    {
        var writer = new OpenXmlWorkbookWriter(stream, styles);
        writer.CreateSheet(items1);
        writer.CreateSheet(items2);
        writer.CreateSheet(items3);
        writer.CreateSheet(items4);
        writer.Save();
    }
    public static void SaveToExcel<T1, T2, T3, T4, T5>(this Stream stream,
        IEnumerable<T1> items1,
        IEnumerable<T2> items2,
        IEnumerable<T3> items3,
        IEnumerable<T4> items4,
        IEnumerable<T5> items5,
        OpenXmlStyles? styles = null)
    {
        var writer = new OpenXmlWorkbookWriter(stream, styles);
        writer.CreateSheet(items1);
        writer.CreateSheet(items2);
        writer.CreateSheet(items3);
        writer.CreateSheet(items4);
        writer.CreateSheet(items5);
        writer.Save();
    }

    public static void SaveToExcel<T1, T2, T3, T4, T5, T6>(this Stream stream,
        IEnumerable<T1> items1,
        IEnumerable<T2> items2,
        IEnumerable<T3> items3,
        IEnumerable<T4> items4,
        IEnumerable<T5> items5,
        IEnumerable<T6> items6,
        OpenXmlStyles? styles = null)
    {
        var writer = new OpenXmlWorkbookWriter(stream, styles);
        writer.CreateSheet(items1);
        writer.CreateSheet(items2);
        writer.CreateSheet(items3);
        writer.CreateSheet(items4);
        writer.CreateSheet(items5);
        writer.CreateSheet(items6);
        writer.Save();
    }
}