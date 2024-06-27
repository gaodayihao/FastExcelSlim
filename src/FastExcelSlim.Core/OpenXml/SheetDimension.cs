using FastExcelSlim.Extensions;
using Utf8StringInterpolation;

namespace FastExcelSlim.OpenXml;

internal class SheetDimension(int maxRow, int maxColumn) : XmlElement
{
    internal int MaxRow { get; } = maxRow;

    internal int MaxColumn { get; } = maxColumn;

    internal override void Write<TBufferWriter>(scoped ref Utf8StringWriter<TBufferWriter> writer)
    {
        writer.AppendLiteral("<dimension ref=\"A1");
        if (MaxRow > 0 && (MaxColumn > 1 || MaxRow > 1))
        {
            writer.AppendLiteral(":");
            writer.ConvertXYToCellReference(MaxColumn, MaxRow);
        }
        writer.AppendLiteral("\"/>");
    }
}