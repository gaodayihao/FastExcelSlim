#if NETSTANDARD2_0
namespace System;

internal static class StringExtensions
{
    public static bool StartsWith(this string source, char value)
    {
        if (source == null) throw new ArgumentNullException(nameof(source));
        return source.AsSpan().StartsWith([value]);
    }

    public static bool EndsWith(this string source, char value)
    {
        if (source == null) throw new ArgumentNullException(nameof(source));
        return source.AsSpan().EndsWith([value]);
    }
}
#endif