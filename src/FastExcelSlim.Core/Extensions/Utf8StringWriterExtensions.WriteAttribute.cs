using System.Buffers;
using System.Drawing;
using Utf8StringInterpolation;

namespace FastExcelSlim.Extensions;

partial class Utf8StringWriterExtensions
{
    public static void WriteAttribute<TBufferWriter, TValue>(this scoped ref Utf8StringWriter<TBufferWriter> writer, string attributeName, TValue value)
        where TValue : Enum
        where TBufferWriter : IBufferWriter<byte>
    {
        writer.WriteAttribute(attributeName, value, "G", true);
    }

    public static void WriteAttribute<TBufferWriter, TValue>(this scoped ref Utf8StringWriter<TBufferWriter> writer, string attributeName, TValue value, string format)
        where TValue : Enum
        where TBufferWriter : IBufferWriter<byte>
    {
        writer.WriteAttribute(attributeName, value, format, true);
    }

    public static void WriteAttribute<TBufferWriter, TValue>(this scoped ref Utf8StringWriter<TBufferWriter> writer, string attributeName, TValue value, string format, bool toLowerCamelCase)
        where TValue : Enum
        where TBufferWriter : IBufferWriter<byte>
    {
        writer.AppendFormat($" {attributeName}");
        writer.AppendLiteral("=\"");
        if ("G".Equals(format, StringComparison.OrdinalIgnoreCase) && toLowerCamelCase)
        {
            var attributeValue = value.ToString(format).AsSpan();
            for (int i = 0; i < attributeValue.Length; i++)
            {
                var c = attributeValue[i];
                if (i == 0 && char.IsUpper(c))
                {
                    writer.Append(char.ToLower(c));
                }
                else writer.Append(c);
            }
        }
        else
        {
            writer.AppendLiteral(value.ToString(format));
        }
        writer.AppendLiteral("\"");
    }

    public static void WriteAttribute<TBufferWriter>(this scoped ref Utf8StringWriter<TBufferWriter> writer, string attributeName, Color color)
        where TBufferWriter : IBufferWriter<byte>
    {
        writer.AppendFormat($" {attributeName}=\"{color.R:X2}{color.G:X2}{color.B:X2}\"");
    }

    public static void WriteAttribute<TBufferWriter>(this scoped ref Utf8StringWriter<TBufferWriter> writer, string attributeName, bool value)
        where TBufferWriter : IBufferWriter<byte>
    {
        writer.WriteAttribute(attributeName, value, false);
    }

    public static void WriteAttribute<TBufferWriter>(this scoped ref Utf8StringWriter<TBufferWriter> writer, string attributeName, bool value, bool writeIfBlank)
        where TBufferWriter : IBufferWriter<byte>
    {
        if (!value && !writeIfBlank) return;
        writer.AppendFormat($" {attributeName}=\"{(value ? "1" : "0")}\"");
    }

    public static void WriteAttribute<TBufferWriter>(this scoped ref Utf8StringWriter<TBufferWriter> writer, string attributeName, long value)
        where TBufferWriter : IBufferWriter<byte>
    {
        writer.WriteAttribute(attributeName, value, false);
    }

    public static void WriteAttribute<TBufferWriter>(this scoped ref Utf8StringWriter<TBufferWriter> writer, string attributeName, long value, bool writeIfBlank)
        where TBufferWriter : IBufferWriter<byte>
    {
        if (value == 0 && !writeIfBlank) return;
        writer.AppendFormat($" {attributeName}=\"{value}\"");
    }

    public static void WriteAttribute<TBufferWriter>(this scoped ref Utf8StringWriter<TBufferWriter> writer, string attributeName, float value)
        where TBufferWriter : IBufferWriter<byte>
    {
        writer.WriteAttribute(attributeName, value, false);
    }

    public static void WriteAttribute<TBufferWriter>(this scoped ref Utf8StringWriter<TBufferWriter> writer, string attributeName, float value, bool writeIfBlank)
        where TBufferWriter : IBufferWriter<byte>
    {
        if (value == 0 && !writeIfBlank) return;
        writer.AppendFormat($" {attributeName}=\"{value}\"");
    }

    public static void WriteAttribute<TBufferWriter>(this scoped ref Utf8StringWriter<TBufferWriter> writer, string attributeName, double value)
        where TBufferWriter : IBufferWriter<byte>
    {
        writer.WriteAttribute(attributeName, value, false);
    }

    public static void WriteAttribute<TBufferWriter>(this scoped ref Utf8StringWriter<TBufferWriter> writer, string attributeName, double value, bool writeIfBlank)
        where TBufferWriter : IBufferWriter<byte>
    {
        if (value == 0 && !writeIfBlank) return;
        writer.AppendFormat($" {attributeName}=\"{value}\"");
    }

    public static void WriteAttribute<TBufferWriter>(this scoped ref Utf8StringWriter<TBufferWriter> writer, string attributeName, int value)
        where TBufferWriter : IBufferWriter<byte>
    {
        writer.WriteAttribute(attributeName, value, false);
    }

    public static void WriteAttribute<TBufferWriter>(this scoped ref Utf8StringWriter<TBufferWriter> writer, string attributeName, int value, bool writeIfBlank)
        where TBufferWriter : IBufferWriter<byte>
    {
        if (value == 0 && !writeIfBlank) return;
        writer.AppendFormat($" {attributeName}=\"{value}\"");
    }

    public static void WriteAttribute<TBufferWriter>(this scoped ref Utf8StringWriter<TBufferWriter> writer, string attributeName, string? value)
        where TBufferWriter : IBufferWriter<byte>
    {
        writer.WriteAttribute(attributeName, value, false);
    }

    public static void WriteAttribute<TBufferWriter>(this scoped ref Utf8StringWriter<TBufferWriter> writer, string attributeName, string? value, bool writeIfBlank)
        where TBufferWriter : IBufferWriter<byte>
    {
        writer.WriteAttribute(attributeName, value, writeIfBlank, string.Empty);
    }

    public static void WriteAttribute<TBufferWriter>(this scoped ref Utf8StringWriter<TBufferWriter> writer, string attributeName, string? value, bool writeIfBlank, string defaultValue)
        where TBufferWriter : IBufferWriter<byte>
    {
        if ((string.IsNullOrEmpty(value) || defaultValue.Equals(value)) && !writeIfBlank) return;

        writer.AppendFormat($" {attributeName}=\"");
        if (value != null)
        {
            writer.EncodeXml(value);
        }
        writer.AppendLiteral("\"");
    }

}