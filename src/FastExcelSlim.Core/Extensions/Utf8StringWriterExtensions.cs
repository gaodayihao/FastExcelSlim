using System.Buffers;
using System.Collections.Immutable;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using Utf8StringInterpolation;

namespace FastExcelSlim.Extensions;

public static partial class Utf8StringWriterExtensions
{
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

    public static void WriteColumn<TBufferWriter>(this scoped ref Utf8StringWriter<TBufferWriter> writer, int columnIndex, float width)
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
        var yDigitCount = CountDigits((uint)row);

        Span<byte> columnChars = stackalloc byte[xDigitCount + yDigitCount];
        var xCharBytes = columnChars[..xDigitCount];
        var yCharBytes = columnChars[xDigitCount..];

        var index = 1;
        var dividend = column;

        while (dividend > 0)
        {
            var modulo = (dividend - 1) % 26;
            xCharBytes[^index++] = (byte)(a + modulo);
            dividend = (dividend - modulo) / 26;
        }

#if NET8_0_OR_GREATER
        row.TryFormat(yCharBytes, out _);
#else
        Span<char> chars = stackalloc char[yDigitCount];
        row.TryFormat(chars, out _);
        for (int i = 0; i < chars.Length; i++)
        {
            yCharBytes[i] = (byte)chars[i];
        }
#endif

        writer.AppendUtf8(columnChars);
    }

    private static readonly char[] XmlEncodeCharsToReplace = ['&', '<', '>', '"'];

    private static readonly IReadOnlyDictionary<char, string> EncodeMap = new Dictionary<char, string>
    {
        { '&', "&amp;" },
        { '<', "&lt;" },
        { '>', "&gt;" },
        { '\"', "&quot;" },
    }.ToImmutableDictionary();

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

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    private static int CountDigits(uint value)
    {
        // Algorithm based on https://lemire.me/blog/2021/06/03/computing-the-number-of-digits-of-an-integer-even-faster.
        ReadOnlySpan<long> table =
        [
            4294967296,
            8589934582,
            8589934582,
            8589934582,
            12884901788,
            12884901788,
            12884901788,
            17179868184,
            17179868184,
            17179868184,
            21474826480,
            21474826480,
            21474826480,
            21474826480,
            25769703776,
            25769703776,
            25769703776,
            30063771072,
            30063771072,
            30063771072,
            34349738368,
            34349738368,
            34349738368,
            34349738368,
            38554705664,
            38554705664,
            38554705664,
            41949672960,
            41949672960,
            41949672960,
            42949672960,
            42949672960,
        ];
        var tableValue = Unsafe.Add(ref MemoryMarshal.GetReference(table), uint.Log2(value));
        return (int)((value + tableValue) >> 32);
    }
}