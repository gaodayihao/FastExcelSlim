
using FastExcelSlim.OpenXml;

namespace FastExcelSlim;

public static class StreamExtensions
{
    public static void WriteAsExcel<T1>(
        this Stream stream,
        IEnumerable<T1> items1,
        OpenXmlStyles? styles = null)
    {
        var writer = new OpenXmlWorkbookWriter(stream, styles);
        writer.CreateSheet(items1);
        writer.Save();
    }

    public static void WriteAsExcel<T1, T2>(
        this Stream stream,
        IEnumerable<T1> items1,
        IEnumerable<T2> items2,
        OpenXmlStyles? styles = null)
    {
        var writer = new OpenXmlWorkbookWriter(stream, styles);
        writer.CreateSheet(items1);
        writer.CreateSheet(items2);
        writer.Save();
    }

    public static void WriteAsExcel<T1, T2, T3>(
        this Stream stream,
        IEnumerable<T1> items1,
        IEnumerable<T2> items2,
        IEnumerable<T3> items3,
        OpenXmlStyles? styles = null)
    {
        var writer = new OpenXmlWorkbookWriter(stream, styles);
        writer.CreateSheet(items1);
        writer.CreateSheet(items2);
        writer.CreateSheet(items3);
        writer.Save();
    }

    public static void WriteAsExcel<T1, T2, T3, T4>(
        this Stream stream,
        IEnumerable<T1> items1,
        IEnumerable<T2> items2,
        IEnumerable<T3> items3,
        IEnumerable<T4> items4,
        OpenXmlStyles? styles = null)
    {
        var writer = new OpenXmlWorkbookWriter(stream, styles);
        writer.CreateSheet(items1);
        writer.CreateSheet(items2);
        writer.CreateSheet(items3);
        writer.CreateSheet(items4);
        writer.Save();
    }

    public static void WriteAsExcel<T1, T2, T3, T4, T5>(
        this Stream stream,
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

    public static void WriteAsExcel<T1, T2, T3, T4, T5, T6>(
        this Stream stream,
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

    public static void WriteAsExcel<T1, T2, T3, T4, T5, T6, T7>(
        this Stream stream,
        IEnumerable<T1> items1,
        IEnumerable<T2> items2,
        IEnumerable<T3> items3,
        IEnumerable<T4> items4,
        IEnumerable<T5> items5,
        IEnumerable<T6> items6,
        IEnumerable<T7> items7,
        OpenXmlStyles? styles = null)
    {
        var writer = new OpenXmlWorkbookWriter(stream, styles);
        writer.CreateSheet(items1);
        writer.CreateSheet(items2);
        writer.CreateSheet(items3);
        writer.CreateSheet(items4);
        writer.CreateSheet(items5);
        writer.CreateSheet(items6);
        writer.CreateSheet(items7);
        writer.Save();
    }

    public static void WriteAsExcel<T1, T2, T3, T4, T5, T6, T7, T8>(
        this Stream stream,
        IEnumerable<T1> items1,
        IEnumerable<T2> items2,
        IEnumerable<T3> items3,
        IEnumerable<T4> items4,
        IEnumerable<T5> items5,
        IEnumerable<T6> items6,
        IEnumerable<T7> items7,
        IEnumerable<T8> items8,
        OpenXmlStyles? styles = null)
    {
        var writer = new OpenXmlWorkbookWriter(stream, styles);
        writer.CreateSheet(items1);
        writer.CreateSheet(items2);
        writer.CreateSheet(items3);
        writer.CreateSheet(items4);
        writer.CreateSheet(items5);
        writer.CreateSheet(items6);
        writer.CreateSheet(items7);
        writer.CreateSheet(items8);
        writer.Save();
    }

    public static void WriteAsExcel<T1, T2, T3, T4, T5, T6, T7, T8, T9>(
        this Stream stream,
        IEnumerable<T1> items1,
        IEnumerable<T2> items2,
        IEnumerable<T3> items3,
        IEnumerable<T4> items4,
        IEnumerable<T5> items5,
        IEnumerable<T6> items6,
        IEnumerable<T7> items7,
        IEnumerable<T8> items8,
        IEnumerable<T9> items9,
        OpenXmlStyles? styles = null)
    {
        var writer = new OpenXmlWorkbookWriter(stream, styles);
        writer.CreateSheet(items1);
        writer.CreateSheet(items2);
        writer.CreateSheet(items3);
        writer.CreateSheet(items4);
        writer.CreateSheet(items5);
        writer.CreateSheet(items6);
        writer.CreateSheet(items7);
        writer.CreateSheet(items8);
        writer.CreateSheet(items9);
        writer.Save();
    }

    public static void WriteAsExcel<T1, T2, T3, T4, T5, T6, T7, T8, T9, T10>(
        this Stream stream,
        IEnumerable<T1> items1,
        IEnumerable<T2> items2,
        IEnumerable<T3> items3,
        IEnumerable<T4> items4,
        IEnumerable<T5> items5,
        IEnumerable<T6> items6,
        IEnumerable<T7> items7,
        IEnumerable<T8> items8,
        IEnumerable<T9> items9,
        IEnumerable<T10> items10,
        OpenXmlStyles? styles = null)
    {
        var writer = new OpenXmlWorkbookWriter(stream, styles);
        writer.CreateSheet(items1);
        writer.CreateSheet(items2);
        writer.CreateSheet(items3);
        writer.CreateSheet(items4);
        writer.CreateSheet(items5);
        writer.CreateSheet(items6);
        writer.CreateSheet(items7);
        writer.CreateSheet(items8);
        writer.CreateSheet(items9);
        writer.CreateSheet(items10);
        writer.Save();
    }

    public static void WriteAsExcel<T1, T2, T3, T4, T5, T6, T7, T8, T9, T10, T11>(
        this Stream stream,
        IEnumerable<T1> items1,
        IEnumerable<T2> items2,
        IEnumerable<T3> items3,
        IEnumerable<T4> items4,
        IEnumerable<T5> items5,
        IEnumerable<T6> items6,
        IEnumerable<T7> items7,
        IEnumerable<T8> items8,
        IEnumerable<T9> items9,
        IEnumerable<T10> items10,
        IEnumerable<T11> items11,
        OpenXmlStyles? styles = null)
    {
        var writer = new OpenXmlWorkbookWriter(stream, styles);
        writer.CreateSheet(items1);
        writer.CreateSheet(items2);
        writer.CreateSheet(items3);
        writer.CreateSheet(items4);
        writer.CreateSheet(items5);
        writer.CreateSheet(items6);
        writer.CreateSheet(items7);
        writer.CreateSheet(items8);
        writer.CreateSheet(items9);
        writer.CreateSheet(items10);
        writer.CreateSheet(items11);
        writer.Save();
    }

    public static void WriteAsExcel<T1, T2, T3, T4, T5, T6, T7, T8, T9, T10, T11, T12>(
        this Stream stream,
        IEnumerable<T1> items1,
        IEnumerable<T2> items2,
        IEnumerable<T3> items3,
        IEnumerable<T4> items4,
        IEnumerable<T5> items5,
        IEnumerable<T6> items6,
        IEnumerable<T7> items7,
        IEnumerable<T8> items8,
        IEnumerable<T9> items9,
        IEnumerable<T10> items10,
        IEnumerable<T11> items11,
        IEnumerable<T12> items12,
        OpenXmlStyles? styles = null)
    {
        var writer = new OpenXmlWorkbookWriter(stream, styles);
        writer.CreateSheet(items1);
        writer.CreateSheet(items2);
        writer.CreateSheet(items3);
        writer.CreateSheet(items4);
        writer.CreateSheet(items5);
        writer.CreateSheet(items6);
        writer.CreateSheet(items7);
        writer.CreateSheet(items8);
        writer.CreateSheet(items9);
        writer.CreateSheet(items10);
        writer.CreateSheet(items11);
        writer.CreateSheet(items12);
        writer.Save();
    }

    public static void WriteAsExcel<T1, T2, T3, T4, T5, T6, T7, T8, T9, T10, T11, T12, T13>(
        this Stream stream,
        IEnumerable<T1> items1,
        IEnumerable<T2> items2,
        IEnumerable<T3> items3,
        IEnumerable<T4> items4,
        IEnumerable<T5> items5,
        IEnumerable<T6> items6,
        IEnumerable<T7> items7,
        IEnumerable<T8> items8,
        IEnumerable<T9> items9,
        IEnumerable<T10> items10,
        IEnumerable<T11> items11,
        IEnumerable<T12> items12,
        IEnumerable<T13> items13,
        OpenXmlStyles? styles = null)
    {
        var writer = new OpenXmlWorkbookWriter(stream, styles);
        writer.CreateSheet(items1);
        writer.CreateSheet(items2);
        writer.CreateSheet(items3);
        writer.CreateSheet(items4);
        writer.CreateSheet(items5);
        writer.CreateSheet(items6);
        writer.CreateSheet(items7);
        writer.CreateSheet(items8);
        writer.CreateSheet(items9);
        writer.CreateSheet(items10);
        writer.CreateSheet(items11);
        writer.CreateSheet(items12);
        writer.CreateSheet(items13);
        writer.Save();
    }

    public static void WriteAsExcel<T1, T2, T3, T4, T5, T6, T7, T8, T9, T10, T11, T12, T13, T14>(
        this Stream stream,
        IEnumerable<T1> items1,
        IEnumerable<T2> items2,
        IEnumerable<T3> items3,
        IEnumerable<T4> items4,
        IEnumerable<T5> items5,
        IEnumerable<T6> items6,
        IEnumerable<T7> items7,
        IEnumerable<T8> items8,
        IEnumerable<T9> items9,
        IEnumerable<T10> items10,
        IEnumerable<T11> items11,
        IEnumerable<T12> items12,
        IEnumerable<T13> items13,
        IEnumerable<T14> items14,
        OpenXmlStyles? styles = null)
    {
        var writer = new OpenXmlWorkbookWriter(stream, styles);
        writer.CreateSheet(items1);
        writer.CreateSheet(items2);
        writer.CreateSheet(items3);
        writer.CreateSheet(items4);
        writer.CreateSheet(items5);
        writer.CreateSheet(items6);
        writer.CreateSheet(items7);
        writer.CreateSheet(items8);
        writer.CreateSheet(items9);
        writer.CreateSheet(items10);
        writer.CreateSheet(items11);
        writer.CreateSheet(items12);
        writer.CreateSheet(items13);
        writer.CreateSheet(items14);
        writer.Save();
    }

    public static void WriteAsExcel<T1, T2, T3, T4, T5, T6, T7, T8, T9, T10, T11, T12, T13, T14, T15>(
        this Stream stream,
        IEnumerable<T1> items1,
        IEnumerable<T2> items2,
        IEnumerable<T3> items3,
        IEnumerable<T4> items4,
        IEnumerable<T5> items5,
        IEnumerable<T6> items6,
        IEnumerable<T7> items7,
        IEnumerable<T8> items8,
        IEnumerable<T9> items9,
        IEnumerable<T10> items10,
        IEnumerable<T11> items11,
        IEnumerable<T12> items12,
        IEnumerable<T13> items13,
        IEnumerable<T14> items14,
        IEnumerable<T15> items15,
        OpenXmlStyles? styles = null)
    {
        var writer = new OpenXmlWorkbookWriter(stream, styles);
        writer.CreateSheet(items1);
        writer.CreateSheet(items2);
        writer.CreateSheet(items3);
        writer.CreateSheet(items4);
        writer.CreateSheet(items5);
        writer.CreateSheet(items6);
        writer.CreateSheet(items7);
        writer.CreateSheet(items8);
        writer.CreateSheet(items9);
        writer.CreateSheet(items10);
        writer.CreateSheet(items11);
        writer.CreateSheet(items12);
        writer.CreateSheet(items13);
        writer.CreateSheet(items14);
        writer.CreateSheet(items15);
        writer.Save();
    }

    public static void WriteAsExcel<T1, T2, T3, T4, T5, T6, T7, T8, T9, T10, T11, T12, T13, T14, T15, T16>(
        this Stream stream,
        IEnumerable<T1> items1,
        IEnumerable<T2> items2,
        IEnumerable<T3> items3,
        IEnumerable<T4> items4,
        IEnumerable<T5> items5,
        IEnumerable<T6> items6,
        IEnumerable<T7> items7,
        IEnumerable<T8> items8,
        IEnumerable<T9> items9,
        IEnumerable<T10> items10,
        IEnumerable<T11> items11,
        IEnumerable<T12> items12,
        IEnumerable<T13> items13,
        IEnumerable<T14> items14,
        IEnumerable<T15> items15,
        IEnumerable<T16> items16,
        OpenXmlStyles? styles = null)
    {
        var writer = new OpenXmlWorkbookWriter(stream, styles);
        writer.CreateSheet(items1);
        writer.CreateSheet(items2);
        writer.CreateSheet(items3);
        writer.CreateSheet(items4);
        writer.CreateSheet(items5);
        writer.CreateSheet(items6);
        writer.CreateSheet(items7);
        writer.CreateSheet(items8);
        writer.CreateSheet(items9);
        writer.CreateSheet(items10);
        writer.CreateSheet(items11);
        writer.CreateSheet(items12);
        writer.CreateSheet(items13);
        writer.CreateSheet(items14);
        writer.CreateSheet(items15);
        writer.CreateSheet(items16);
        writer.Save();
    }

    public static void WriteAsExcel<T1, T2, T3, T4, T5, T6, T7, T8, T9, T10, T11, T12, T13, T14, T15, T16, T17>(
        this Stream stream,
        IEnumerable<T1> items1,
        IEnumerable<T2> items2,
        IEnumerable<T3> items3,
        IEnumerable<T4> items4,
        IEnumerable<T5> items5,
        IEnumerable<T6> items6,
        IEnumerable<T7> items7,
        IEnumerable<T8> items8,
        IEnumerable<T9> items9,
        IEnumerable<T10> items10,
        IEnumerable<T11> items11,
        IEnumerable<T12> items12,
        IEnumerable<T13> items13,
        IEnumerable<T14> items14,
        IEnumerable<T15> items15,
        IEnumerable<T16> items16,
        IEnumerable<T17> items17,
        OpenXmlStyles? styles = null)
    {
        var writer = new OpenXmlWorkbookWriter(stream, styles);
        writer.CreateSheet(items1);
        writer.CreateSheet(items2);
        writer.CreateSheet(items3);
        writer.CreateSheet(items4);
        writer.CreateSheet(items5);
        writer.CreateSheet(items6);
        writer.CreateSheet(items7);
        writer.CreateSheet(items8);
        writer.CreateSheet(items9);
        writer.CreateSheet(items10);
        writer.CreateSheet(items11);
        writer.CreateSheet(items12);
        writer.CreateSheet(items13);
        writer.CreateSheet(items14);
        writer.CreateSheet(items15);
        writer.CreateSheet(items16);
        writer.CreateSheet(items17);
        writer.Save();
    }

    public static void WriteAsExcel<T1, T2, T3, T4, T5, T6, T7, T8, T9, T10, T11, T12, T13, T14, T15, T16, T17, T18>(
        this Stream stream,
        IEnumerable<T1> items1,
        IEnumerable<T2> items2,
        IEnumerable<T3> items3,
        IEnumerable<T4> items4,
        IEnumerable<T5> items5,
        IEnumerable<T6> items6,
        IEnumerable<T7> items7,
        IEnumerable<T8> items8,
        IEnumerable<T9> items9,
        IEnumerable<T10> items10,
        IEnumerable<T11> items11,
        IEnumerable<T12> items12,
        IEnumerable<T13> items13,
        IEnumerable<T14> items14,
        IEnumerable<T15> items15,
        IEnumerable<T16> items16,
        IEnumerable<T17> items17,
        IEnumerable<T18> items18,
        OpenXmlStyles? styles = null)
    {
        var writer = new OpenXmlWorkbookWriter(stream, styles);
        writer.CreateSheet(items1);
        writer.CreateSheet(items2);
        writer.CreateSheet(items3);
        writer.CreateSheet(items4);
        writer.CreateSheet(items5);
        writer.CreateSheet(items6);
        writer.CreateSheet(items7);
        writer.CreateSheet(items8);
        writer.CreateSheet(items9);
        writer.CreateSheet(items10);
        writer.CreateSheet(items11);
        writer.CreateSheet(items12);
        writer.CreateSheet(items13);
        writer.CreateSheet(items14);
        writer.CreateSheet(items15);
        writer.CreateSheet(items16);
        writer.CreateSheet(items17);
        writer.CreateSheet(items18);
        writer.Save();
    }

    public static void WriteAsExcel<T1, T2, T3, T4, T5, T6, T7, T8, T9, T10, T11, T12, T13, T14, T15, T16, T17, T18, T19>(
        this Stream stream,
        IEnumerable<T1> items1,
        IEnumerable<T2> items2,
        IEnumerable<T3> items3,
        IEnumerable<T4> items4,
        IEnumerable<T5> items5,
        IEnumerable<T6> items6,
        IEnumerable<T7> items7,
        IEnumerable<T8> items8,
        IEnumerable<T9> items9,
        IEnumerable<T10> items10,
        IEnumerable<T11> items11,
        IEnumerable<T12> items12,
        IEnumerable<T13> items13,
        IEnumerable<T14> items14,
        IEnumerable<T15> items15,
        IEnumerable<T16> items16,
        IEnumerable<T17> items17,
        IEnumerable<T18> items18,
        IEnumerable<T19> items19,
        OpenXmlStyles? styles = null)
    {
        var writer = new OpenXmlWorkbookWriter(stream, styles);
        writer.CreateSheet(items1);
        writer.CreateSheet(items2);
        writer.CreateSheet(items3);
        writer.CreateSheet(items4);
        writer.CreateSheet(items5);
        writer.CreateSheet(items6);
        writer.CreateSheet(items7);
        writer.CreateSheet(items8);
        writer.CreateSheet(items9);
        writer.CreateSheet(items10);
        writer.CreateSheet(items11);
        writer.CreateSheet(items12);
        writer.CreateSheet(items13);
        writer.CreateSheet(items14);
        writer.CreateSheet(items15);
        writer.CreateSheet(items16);
        writer.CreateSheet(items17);
        writer.CreateSheet(items18);
        writer.CreateSheet(items19);
        writer.Save();
    }

    public static void WriteAsExcel<T1, T2, T3, T4, T5, T6, T7, T8, T9, T10, T11, T12, T13, T14, T15, T16, T17, T18, T19, T20>(
        this Stream stream,
        IEnumerable<T1> items1,
        IEnumerable<T2> items2,
        IEnumerable<T3> items3,
        IEnumerable<T4> items4,
        IEnumerable<T5> items5,
        IEnumerable<T6> items6,
        IEnumerable<T7> items7,
        IEnumerable<T8> items8,
        IEnumerable<T9> items9,
        IEnumerable<T10> items10,
        IEnumerable<T11> items11,
        IEnumerable<T12> items12,
        IEnumerable<T13> items13,
        IEnumerable<T14> items14,
        IEnumerable<T15> items15,
        IEnumerable<T16> items16,
        IEnumerable<T17> items17,
        IEnumerable<T18> items18,
        IEnumerable<T19> items19,
        IEnumerable<T20> items20,
        OpenXmlStyles? styles = null)
    {
        var writer = new OpenXmlWorkbookWriter(stream, styles);
        writer.CreateSheet(items1);
        writer.CreateSheet(items2);
        writer.CreateSheet(items3);
        writer.CreateSheet(items4);
        writer.CreateSheet(items5);
        writer.CreateSheet(items6);
        writer.CreateSheet(items7);
        writer.CreateSheet(items8);
        writer.CreateSheet(items9);
        writer.CreateSheet(items10);
        writer.CreateSheet(items11);
        writer.CreateSheet(items12);
        writer.CreateSheet(items13);
        writer.CreateSheet(items14);
        writer.CreateSheet(items15);
        writer.CreateSheet(items16);
        writer.CreateSheet(items17);
        writer.CreateSheet(items18);
        writer.CreateSheet(items19);
        writer.CreateSheet(items20);
        writer.Save();
    }

}