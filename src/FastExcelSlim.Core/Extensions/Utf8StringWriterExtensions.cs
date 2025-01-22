using System.Buffers;
using System.Diagnostics;
using System.Runtime.CompilerServices;
using Utf8StringInterpolation;

#if NET7_0_OR_GREATER
using System.Collections.Immutable;
using System.Runtime.InteropServices;
#endif

namespace FastExcelSlim.Extensions;

public static partial class Utf8StringWriterExtensions
{
    public static void EncodeXml<TBufferWriter>(this scoped ref Utf8StringWriter<TBufferWriter> writer, char? value)
        where TBufferWriter : IBufferWriter<byte>
    {
        if (!value.HasValue) return;
        if (EncodeMap.TryGetValue(value.Value, out var encode))
        {
            writer.AppendLiteral(encode);
        }
        else
        {
            writer.Append(value.Value);
        }
    }

    public static void EncodeXml<TBufferWriter>(this scoped ref Utf8StringWriter<TBufferWriter> writer, string? xml) where TBufferWriter : IBufferWriter<byte>
    {
        if (xml == null) return;

        // quick check whether needed
        if (xml.IndexOfAny(XmlEncodeCharsToReplace) == -1)
        {
            writer.AppendLiteral(xml);
            return;
        }

        foreach (var c in xml.AsSpan())
        {
            if (EncodeMap.TryGetValue(c, out var encode))
            {
                writer.AppendLiteral(encode);
            }
            else
            {
                writer.Append(c);
            }
        }
    }

    public static void WriteHeader<TBufferWriter>(this scoped ref Utf8StringWriter<TBufferWriter> writer,
        OpenXmlStyles styles,
        string headerName,
        string propertyName,
        int columnIndex)
        where TBufferWriter : IBufferWriter<byte>
    {
        const int headerRow = 1;
        var styleIndex = styles.GetHeaderStyleIndex(propertyName);
        var preserveSpace = !string.IsNullOrEmpty(headerName) && (headerName.StartsWith(' ') || headerName.EndsWith(' '));
        writer.AppendLiteral("<c r=\"");
        writer.ConvertXYToCellReference(columnIndex, headerRow);
        writer.AppendFormat($"\" t=\"str\" s=\"{styleIndex}\"{(preserveSpace ? " xml:space=\"preserve\"" : string.Empty)}><v>");
        writer.EncodeXml(headerName);
        writer.AppendLiteral("</v></c>");
    }

    public static void WriteColumn<TBufferWriter>(this scoped ref Utf8StringWriter<TBufferWriter> writer, int columnIndex, double width)
        where TBufferWriter : IBufferWriter<byte>
    {
        writer.AppendFormat($"<col min=\"{columnIndex}\" max=\"{columnIndex}\" width=\"{width}\" customWidth=\"1\" collapsed=\"1\"/>");
    }

    public static void ConvertXYToCellReference<TBufferWriter>(this scoped ref Utf8StringWriter<TBufferWriter> writer, int column, int row) where TBufferWriter : IBufferWriter<byte>
    {
#if NET8_0_OR_GREATER
        ArgumentOutOfRangeException.ThrowIfLessThanOrEqual(column, 0, nameof(column));
        ArgumentOutOfRangeException.ThrowIfLessThanOrEqual(row, 0, nameof(row));
#else
        if (column <= 0)
        {
            throw new ArgumentOutOfRangeException(nameof(column));
        }
        if (row <= 0)
        {
            throw new ArgumentOutOfRangeException(nameof(row));
        }
#endif

        const byte a = (byte)'A';
        var xDigitCount = CountXDigits((uint)column);
        var yDigitCount = FormattingHelpers.CountDigits((uint)row);

        Span<byte> columnChars = stackalloc byte[xDigitCount + yDigitCount];
        var xCharBytes = columnChars.Slice(0,xDigitCount);
        var yCharBytes = columnChars.Slice(xDigitCount);

        var index = 1;
        var dividend = column;

        while (dividend > 0)
        {
            var modulo = (dividend - 1) % 26;
            xCharBytes[xCharBytes.Length - index++] = (byte)(a + modulo);
            dividend = (dividend - modulo) / 26;
        }

        row.TryFormat(yCharBytes, out _);

        writer.AppendUtf8(columnChars);
    }

    private static readonly char[] XmlEncodeCharsToReplace = ['&', '<', '>', '"'];

    private static readonly IReadOnlyDictionary<char, string> EncodeMap = new Dictionary<char, string>
    {
        { '&', "&amp;" },
        { '<', "&lt;" },
        { '>', "&gt;" },
        { '\"', "&quot;" },
#if NET7_0_OR_GREATER
    }.ToImmutableDictionary();
#else
    };
#endif

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    private static int CountXDigits(uint x)
    {
        var digitCount = 0;

        var dividend = x;
        while (dividend > 0)
        {
            digitCount++;
            var modulo = (dividend - 1) % 26;
            dividend = (dividend - modulo) / 26;
        }

        return digitCount;
    }
}