using System.Buffers;
using System.Collections;
using Utf8StringInterpolation;

namespace FastExcelSlim.OpenXml;

public abstract class XmlElement
{
    internal abstract void Write<TBufferWriter>(scoped ref Utf8StringWriter<TBufferWriter> writer) where TBufferWriter : IBufferWriter<byte>;
}

public abstract class IdXmlElement : XmlElement
{
    internal IdXmlElement(int id)
    {
        Id = id;
    }

    public int Id { get; }
}

public abstract class XmlElementCollection<T> : IEnumerable<T> where T : XmlElement
{
    protected List<T> InnerElements = new();

    internal abstract string Name { get; }

    public IEnumerator<T> GetEnumerator()
    {
        return InnerElements.GetEnumerator();
    }

    IEnumerator IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }

    protected virtual void Add(T item)
    {
        InnerElements.Add(item);
    }

    public int Count => InnerElements.Count;

    internal virtual void Write<TBufferWriter>(scoped ref Utf8StringWriter<TBufferWriter> writer)
        where TBufferWriter : IBufferWriter<byte>
    {
        if (Count == 0) return;
        writer.AppendFormat($"<{Name} count=\"{Count}\">");
        foreach (var element in InnerElements)
        {
            element.Write(ref writer);
        }
        writer.AppendFormat($"</{Name}>");
    }
}