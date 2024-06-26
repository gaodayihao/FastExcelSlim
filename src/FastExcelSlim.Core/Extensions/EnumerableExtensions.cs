using FastExcelSlim.OpenXml;

namespace FastExcelSlim;

public static class EnumerableExtensions
{
    public static void SaveAs<T>(this IEnumerable<T> values, Stream stream, OpenXmlStyles? styles = null)
    {
        var writer = new OpenXmlWorkbookWriter(stream, styles);
        writer.CreateSheet(values);
        writer.Save();
    }

    public static void SaveAs<T>(this IEnumerable<T> values, string path, OpenXmlStyles? styles = null)
    {
        using var stream = File.Open(path, FileMode.Create, FileAccess.ReadWrite, FileShare.ReadWrite);
        var writer = new OpenXmlWorkbookWriter(stream, styles);
        writer.CreateSheet(values);
        writer.Save();
    }
}