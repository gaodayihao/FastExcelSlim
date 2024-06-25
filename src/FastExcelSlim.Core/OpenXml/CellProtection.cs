using FastExcelSlim.Extensions;
using Utf8StringInterpolation;

namespace FastExcelSlim.OpenXml;

public class CellProtection : XmlElement
{
    internal CellProtection() { }

    public bool IsLocked { get; set; }

    public bool IsHidden { get; set; }

    internal override void Write<TBufferWriter>(scoped ref Utf8StringWriter<TBufferWriter> writer)
    {
        writer.AppendLiteral("<protection");
        writer.WriteAttribute("locked", IsLocked, true);
        writer.WriteAttribute("hidden", IsHidden);
        writer.AppendLiteral("/>");
    }
}