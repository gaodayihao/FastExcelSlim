using FastExcelSlim.Extensions;
using Utf8StringInterpolation;

namespace FastExcelSlim.OpenXml;

internal class SheetDimension(int maxRow, int maxColumn) : XmlElement
{
    internal override void Write<TBufferWriter>(scoped ref Utf8StringWriter<TBufferWriter> writer)
    {
        writer.AppendLiteral("<dimension ref=\"A1");
        if (maxRow > 0 && (maxColumn > 1 || maxRow > 1))
        {
            writer.AppendLiteral(":");
            writer.ConvertXYToCellReference(maxColumn, maxRow);
        }
        writer.AppendLiteral("\"/>");
    }
}