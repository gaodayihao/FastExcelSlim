#if NET7_0_OR_GREATER
using System.Buffers;
using Utf8StringInterpolation;
#endif

namespace FastExcelSlim;

public interface IOpenXmlWritable<T> where T : IOpenXmlWritable<T>
{
#if NET7_0_OR_GREATER
    static abstract int ColumnCount { get; }

    static abstract string? SheetName { get; }

    static abstract void RegisterFormatter();

    static abstract void WriteCell<TBufferWriter>(scoped ref Utf8StringWriter<TBufferWriter> writer, OpenXmlStyles styles, int rowIndex, scoped ref T value) where TBufferWriter : IBufferWriter<byte>;

    static abstract void WriteColumns<TBufferWriter>(scoped ref Utf8StringWriter<TBufferWriter> writer) where TBufferWriter : IBufferWriter<byte>;

    static abstract void WriteHeaders<TBufferWriter>(scoped ref Utf8StringWriter<TBufferWriter> writer, OpenXmlStyles styles) where TBufferWriter : IBufferWriter<byte>;

    static abstract bool FreezeHeader { get; }

    static abstract bool AutoFilter { get; }
#endif
}