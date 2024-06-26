using System.Buffers;
using System.Drawing;
using FastExcelSlim.Extensions;
using Utf8StringInterpolation;

namespace FastExcelSlim.OpenXml;

public class OpenXmlColor : XmlElement
{
    internal OpenXmlColor(int? colorIndex)
    {
        ColorIndex = colorIndex;
    }

    internal OpenXmlColor(Color? color)
    {
        Color = color;
    }

    public int? ColorIndex { get; }

    public Color? Color { get; }

    internal void Write<TBufferWriter>(scoped ref Utf8StringWriter<TBufferWriter> writer, string nodeName)
        where TBufferWriter : IBufferWriter<byte>
    {
        writer.AppendFormat($"<{nodeName}");
        if (ColorIndex.HasValue)
        {
            writer.WriteAttribute("indexed", ColorIndex.Value);
        }
        if (Color.HasValue)
        {
            writer.WriteAttribute("rgb", Color.Value);
        }
        writer.AppendLiteral("/>");
    }

    internal override void Write<TBufferWriter>(scoped ref Utf8StringWriter<TBufferWriter> writer)
    {
        writer.AppendLiteral("<color");
        if (ColorIndex.HasValue)
        {
            writer.WriteAttribute("indexed", ColorIndex.Value);
        }
        if (Color.HasValue)
        {
            writer.WriteAttribute("rgb", Color.Value);
        }
        writer.AppendLiteral("/>");
    }
}