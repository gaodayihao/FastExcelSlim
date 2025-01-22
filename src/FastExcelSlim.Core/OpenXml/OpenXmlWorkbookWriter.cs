using FastExcelSlim.BufferWriter;
using System.IO.Compression;
using System.IO.Pipelines;
using System.Text;
using Utf8StringInterpolation;

namespace FastExcelSlim.OpenXml;

public class OpenXmlWorkbookWriter(Stream stream, OpenXmlStyles? styles)
{
    protected static readonly UTF8Encoding UTF8WithBom = new(true);
    private readonly ZipArchive _archive = new(stream, ZipArchiveMode.Update, true, UTF8WithBom);
    private readonly OpenXmlStyles _styles = styles ?? DefaultStyles.Default;
    private readonly Dictionary<string, string> _zipEntries = new();
    private readonly OpenXmlWorkbook _workbook = new();

    public OpenXmlWorkbookWriter(Stream stream) : this(stream, null)
    {
    }

    public void CreateSheet<T>(IEnumerable<T> values)
    {
        _workbook.CreateSheet(values, _styles);
    }

    public void Save()
    {
        GenerateDefaultOpenXml();
        GenerateWorkbookOpenXml();
        GenerateContentTypes();
        _archive.Dispose();
    }

    private void CreateZipEntry(string path, string? content, string? contentType = null)
    {
        var entry = _archive.CreateEntry(path, CompressionLevel.Fastest);
        using var zipStream = entry.Open();
        var pipeWriter = new StreamPipeWriter(zipStream, new StreamPipeWriterOptions(leaveOpen: true));
        using var writer = Utf8String.CreateWriter(pipeWriter);
        writer.Append(content);
        writer.Flush();
        pipeWriter.Complete();
        if (!string.IsNullOrEmpty(contentType))
        {
            _zipEntries[path] = contentType!;
        }
    }

    private ZipEntryWriterWrapper CreateZipEntryWriterWrapper(string path, string? contentType = null)
    {
        var entry = _archive.CreateEntry(path, CompressionLevel.Fastest);
        var zipStream = entry.Open();
        var pipeWriter = new StreamPipeWriter(zipStream, new StreamPipeWriterOptions(leaveOpen: true));
        if (!string.IsNullOrEmpty(contentType))
        {
            _zipEntries[path] = contentType!;
        }

        return new ZipEntryWriterWrapper(zipStream, pipeWriter);
    }

    internal void GenerateDefaultOpenXml()
    {
        CreateZipEntry(ExcelFilePath.Rels, ExcelXml.DefaultRels);

        CreateZipEntry(ExcelFilePath.ExtendedProperties, ExcelXml.ExtendedProperties, ExcelContentTypes.ExtendedProperties);

        var wrapper = CreateZipEntryWriterWrapper(ExcelFilePath.CustomProperties, ExcelContentTypes.CustomProperties);
        wrapper.Writer.AppendFormat($"{ExcelXml.CustomPropertiesStart}{typeof(OpenXmlWorkbookWriter).Assembly.GetName().Version!.ToString(3)}{ExcelXml.CustomPropertiesEnd}");
        wrapper.Dispose();

        wrapper = CreateZipEntryWriterWrapper(ExcelFilePath.CoreProperties, ExcelContentTypes.CoreProperties);
        wrapper.Writer.AppendFormat($"{ExcelXml.CorePropertiesStart}{DateTime.UtcNow:yyyy'-'MM'-'dd'T'HH':'mm':'ss'Z'}{ExcelXml.CorePropertiesEnd}");
        wrapper.Dispose();

        CreateZipEntry(ExcelFilePath.SharedStrings, ExcelXml.DefaultSharedStrings, ExcelContentTypes.SharedStrings);

        wrapper = CreateZipEntryWriterWrapper(ExcelFilePath.Styles, ExcelContentTypes.Styles);
        _styles.Write(ref wrapper.Writer);
        wrapper.Dispose();
    }

    internal void GenerateWorkbookOpenXml()
    {
        {
            var wrapper = CreateZipEntryWriterWrapper(ExcelFilePath.WorkbookRelationships);
            _workbook.WriteRelationships(ref wrapper);
            wrapper.Dispose();

            wrapper = CreateZipEntryWriterWrapper(ExcelFilePath.Workbook, ExcelContentTypes.Workbook);
            _workbook.Write(ref wrapper);
            wrapper.Dispose();
        }

        foreach (var sheet in _workbook.Sheets)
        {
            var wrapper = CreateZipEntryWriterWrapper(sheet.Path, ExcelContentTypes.Worksheet);
            sheet.Write(ref wrapper);
            wrapper.Dispose();
        }
    }

    internal void GenerateContentTypes()
    {
        var wrapper = CreateZipEntryWriterWrapper(ExcelFilePath.ContentTypes);
        wrapper.Writer.AppendFormat($"{ExcelXml.ContentTypesStart}{ExcelXml.ContentTypesDefault}");
#if !NETSTANDARD2_0
        foreach (var (path, type) in _zipEntries)
        {
#else
        foreach (var kvp in _zipEntries)
        {
            var path = kvp.Key;
            var type = kvp.Value;
#endif
            wrapper.Writer.AppendFormat($"<Override PartName=\"/{path}\" ContentType=\"{type}\" />");
        }
        wrapper.Writer.AppendLiteral(ExcelXml.ContentTypesEnd);
        wrapper.Dispose();
    }
}

internal ref struct ZipEntryWriterWrapper(Stream stream, StreamPipeWriter pipeWriter)
{
    //private static Action<PipeWriter, bool>? _internalFlush;

    //private static void PipeFlush(PipeWriter writer)
    //{
    //    _internalFlush ??= CreateDelegate(writer);
    //    _internalFlush?.Invoke(writer, true);
    //}

    //private static Action<PipeWriter, bool> CreateDelegate(PipeWriter writer)
    //{
    //    var method = new DynamicMethod("CallFlushInternal", typeof(void), [typeof(PipeWriter), typeof(bool)]);
    //    var il = method.GetILGenerator();
    //    il.Emit(OpCodes.Ldarg_0);
    //    il.Emit(OpCodes.Ldarg_1);
    //    il.Emit(OpCodes.Callvirt, writer.GetType().GetMethod("FlushInternal", BindingFlags.NonPublic | BindingFlags.Instance)!);
    //    il.Emit(OpCodes.Ret);

    //    return method.CreateDelegate<Action<PipeWriter, bool>>();
    //}

    private const int BufferSize = 64 * 1024;

    public Utf8StringWriter<PipeWriter> Writer = Utf8String.CreateWriter((PipeWriter)pipeWriter);

    public void CheckAndFlush()
    {
        if (pipeWriter.UnflushedBytes >= BufferSize)
        {
            Writer.ClearState();
            pipeWriter.FlushInternal(true);
        }
    }

    public void Dispose()
    {
        Writer.Flush();
        pipeWriter.Complete();
        Writer.Dispose();
        stream.Dispose();
    }
}