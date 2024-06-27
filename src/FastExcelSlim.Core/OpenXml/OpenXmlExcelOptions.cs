namespace FastExcelSlim.OpenXml;

public ref struct OpenXmlExcelOptions(bool freezeHeader, bool autoFilter)
{
    public bool FreezeHeader = freezeHeader;

    public bool AutoFilter = autoFilter;
}