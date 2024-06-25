using FastExcelSlim.OpenXml;

namespace FastExcelSlim;

public static class EnumerableExtensions
{
    public static void SaveAs<T>(this IEnumerable<T> values, Stream stream) where T : IOpenXmlWritable<T>
    {
        new OpenXmlWorkbookWriter<T>(stream, values).Save();
    }

    public static void SaveAs<T>(this IEnumerable<T> values, string path) where T : IOpenXmlWritable<T>
    {
        using var stream = File.Open(path, FileMode.Create, FileAccess.ReadWrite, FileShare.ReadWrite);
        new OpenXmlWorkbookWriter<T>(stream, values).Save();
    }
}