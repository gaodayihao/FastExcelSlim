using System.Diagnostics;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace FastExcelSlim;

internal class FormattingHelpers
{
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static int CountDigits(uint value)
    {
#if NET7_0_OR_GREATER
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
#else
        int digits = 1;
        if (value >= 100000)
        {
            value /= 100000;
            digits += 5;
        }

        if (value < 10)
        {
            // no-op
        }
        else if (value < 100)
        {
            digits += 1;
        }
        else if (value < 1000)
        {
            digits += 2;
        }
        else if (value < 10000)
        {
            digits += 3;
        }
        else
        {
            Debug.Assert(value < 100000);
            digits += 4;
        }

        return digits;
#endif
    }
}