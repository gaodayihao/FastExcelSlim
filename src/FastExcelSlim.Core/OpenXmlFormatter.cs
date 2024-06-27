using System.Buffers;
using FastExcelSlim.OpenXml;
using Utf8StringInterpolation;

namespace FastExcelSlim;

public interface IOpenXmlFormatter<T>
{
    void WriteCell<TBufferWriter>(scoped ref Utf8StringWriter<TBufferWriter> writer, OpenXmlStyles styles, int rowIndex, scoped ref T value) where TBufferWriter : IBufferWriter<byte>;

    void WriteColumns<TBufferWriter>(scoped ref Utf8StringWriter<TBufferWriter> writer) where TBufferWriter : IBufferWriter<byte>;

    void WriteHeaders<TBufferWriter>(scoped ref Utf8StringWriter<TBufferWriter> writer, OpenXmlStyles styles) where TBufferWriter : IBufferWriter<byte>;

    int ColumnCount { get; }

    string? SheetName { get; }

    OpenXmlExcelOptions GetOptions();
}

#if NET7_0_OR_GREATER
public sealed class OpenXmlFormatter<T> : IOpenXmlFormatter<T> where T : IOpenXmlWritable<T>
{
    public void WriteCell<TBufferWriter>(scoped ref Utf8StringWriter<TBufferWriter> writer, OpenXmlStyles styles, int rowIndex, scoped ref T value) where TBufferWriter : IBufferWriter<byte>
    {
        T.WriteCell(ref writer, styles, rowIndex, ref value);
    }

    public void WriteColumns<TBufferWriter>(scoped ref Utf8StringWriter<TBufferWriter> writer) where TBufferWriter : IBufferWriter<byte>
    {
        T.WriteColumns(ref writer);
    }

    public void WriteHeaders<TBufferWriter>(scoped ref Utf8StringWriter<TBufferWriter> writer, OpenXmlStyles styles) where TBufferWriter : IBufferWriter<byte>
    {
        T.WriteHeaders(ref writer, styles);
    }

    public int ColumnCount => T.ColumnCount;

    public string? SheetName => T.SheetName;

    public OpenXmlExcelOptions GetOptions()
    {
        return T.GetOptions();
    }
}
#endif