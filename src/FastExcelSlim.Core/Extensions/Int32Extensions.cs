using FastExcelSlim;
using System.Diagnostics;
using System.Runtime.CompilerServices;

namespace System;

internal static class Int32Extensions
{
#if !NET8_0_OR_GREATER
    public static bool TryFormat(this int source, Span<byte> destination, out int charsWritten)
    {
        Span<char> chars = stackalloc char[destination.Length];
        var canWrite = source.TryFormat(chars, out charsWritten);
        if (canWrite)
        {
            for (int i = 0; i < chars.Length; i++)
            {
                destination[i] = (byte)chars[i];
            }
        }
        return canWrite;
    }
#endif

#if NETSTANDARD2_0
    public static bool TryFormat(this int value, Span<char> destination, out int charsWritten)
    {
        if (value < 0) throw new ArgumentOutOfRangeException(nameof(value));

        int bufferLength = FormattingHelpers.CountDigits((uint)value);

        if (bufferLength > destination.Length)
        {
            charsWritten = 0;
            return false;
        }

        charsWritten = bufferLength;
        var uintValue = (uint)value;
        var offset = destination.Length;
        while (--bufferLength >= 0 || uintValue != 0)
        {
            uint newValue = uintValue / 10;
            destination[--offset] = (char)(uintValue - newValue * 10 + '0');
            uintValue = newValue;
        }
        return true;
    }
#endif
}