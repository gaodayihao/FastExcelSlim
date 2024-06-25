using System.Data;

namespace FastExcelSlim.OpenXml;

internal static class ExcelXml
{
    internal const string DefaultSheetName = "OpenXml_Default_Sheet_Name";

    internal const string DefaultRels = @"<?xml version=""1.0"" encoding=""utf-8""?>
<Relationships xmlns=""http://schemas.openxmlformats.org/package/2006/relationships"">
    <Relationship Id=""rId1"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"" Target=""xl/workbook.xml"" />
    <Relationship Id=""rId2"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties"" Target=""docProps/app.xml"" />
    <Relationship Id=""rId3"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties"" Target=""docProps/custom.xml"" />
    <Relationship Id=""rId4"" Type=""http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties"" Target=""docProps/core.xml"" />
</Relationships>";

    internal const string ExtendedProperties = @"<?xml version=""1.0"" encoding=""utf-8""?>
<Properties xmlns:vt=""http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"" xmlns=""http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"">
  <ScaleCrop>false</ScaleCrop>
  <LinksUpToDate>false</LinksUpToDate>
  <SharedDoc>false</SharedDoc>
  <HyperlinksChanged>false</HyperlinksChanged>
  <Application>FastExcelSlim</Application>
  <DocSecurity>0</DocSecurity>
</Properties>";

    internal const string CustomPropertiesStart = @"<?xml version=""1.0"" encoding=""utf-8""?>
<Properties xmlns=""http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"">
  <property fmtid=""{D5CDD505-2E9C-101B-9397-08002B2CF9AE}"" pid=""2"" name=""Generator"">
    <lpwstr xmlns=""http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"">FastExcelSlim</lpwstr>
  </property>
  <property fmtid=""{D5CDD505-2E9C-101B-9397-08002B2CF9AE}"" pid=""3"" name=""Generator Version"">
    <lpwstr xmlns=""http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"">";

    internal const string CustomPropertiesEnd = @"</lpwstr>
  </property>
</Properties>";

    internal const string CorePropertiesStart = @"<?xml version=""1.0"" encoding=""utf-8""?>
<coreProperties xmlns:cp=""http://schemas.openxmlformats.org/package/2006/metadata/core-properties""
    xmlns:dc=""http://purl.org/dc/elements/1.1/""
    xmlns:dcterms=""http://purl.org/dc/terms/""
    xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance""
    xmlns=""http://schemas.openxmlformats.org/package/2006/metadata/core-properties"">
    <dcterms:created xsi:type=""dcterms:W3CDTF"">";

    internal const string CorePropertiesEnd = @"</dcterms:created>
    <dc:creator>FastExcelSlim</dc:creator>
</coreProperties>";

    internal const string DefaultSharedStrings = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\" ?><sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"0\" uniqueCount=\"0\"></sst>";

    internal const string StylesStart = @"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
<styleSheet xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main""
    xmlns:mc=""http://schemas.openxmlformats.org/markup-compatibility/2006"" mc:Ignorable=""x14ac x16r2 xr""
    xmlns:x14ac=""http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac""
    xmlns:x16r2=""http://schemas.microsoft.com/office/spreadsheetml/2015/02/main""
    xmlns:xr=""http://schemas.microsoft.com/office/spreadsheetml/2014/revision"">
";

    internal const string StylesEnd = @"
</styleSheet>";

    internal const string SheetStart = @"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
<worksheet xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main""
    xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships""
    xmlns:mc=""http://schemas.openxmlformats.org/markup-compatibility/2006"" mc:Ignorable=""x14ac xr xr2 xr3""
    xmlns:x14ac=""http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac""
    xmlns:xr=""http://schemas.microsoft.com/office/spreadsheetml/2014/revision""
    xmlns:xr2=""http://schemas.microsoft.com/office/spreadsheetml/2015/revision2""
    xmlns:xr3=""http://schemas.microsoft.com/office/spreadsheetml/2016/revision3"">
";

    internal const string SheetViews = @"<sheetViews>
        <sheetView tabSelected=""1"" showRuler=""1"" showOutlineSymbols=""1"" defaultGridColor=""1"" colorId=""64"" zoomScale=""100"" workbookViewId=""0""></sheetView>
    </sheetViews>
";

    internal const string SheetFormatPr = @"<sheetFormatPr defaultColWidth=""8.43"" defaultRowHeight=""15""/>
";

    internal const string SheetPageMargins = @"<pageMargins left=""0.7"" right=""0.7"" top=""0.75"" bottom=""0.75"" header=""0.3"" footer=""0.3""/>
";

    internal const string SheetEnd = "</worksheet>";

    internal const string WorkbookRelationshipStart = @"<?xml version=""1.0"" encoding=""utf-8""?>
<Relationships xmlns=""http://schemas.openxmlformats.org/package/2006/relationships"">
";

    internal const string WorkbookRelationshipEnd = "</Relationships>";

    internal const string WorkbookStart = @"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
<workbook xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main""
    xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships""
    xmlns:mc=""http://schemas.openxmlformats.org/markup-compatibility/2006"" mc:Ignorable=""x15 xr xr6 xr10 xr2""
    xmlns:x15=""http://schemas.microsoft.com/office/spreadsheetml/2010/11/main""
    xmlns:xr=""http://schemas.microsoft.com/office/spreadsheetml/2014/revision""
    xmlns:xr6=""http://schemas.microsoft.com/office/spreadsheetml/2016/revision6""
    xmlns:xr10=""http://schemas.microsoft.com/office/spreadsheetml/2016/revision10""
    xmlns:xr2=""http://schemas.microsoft.com/office/spreadsheetml/2015/revision2"">
";

    internal const string WorkbookEnd = "</workbook>";

    internal const string WorkbookPr = @"<workbookPr autoCompressPictures=""1""/>";
    internal const string WorkbookViews = @"<bookViews>
        <workbookView tabRatio=""600""/>
    </bookViews>
";

    internal const string ContentTypesStart = @"<?xml version=""1.0"" encoding=""utf-8""?>
<Types xmlns=""http://schemas.openxmlformats.org/package/2006/content-types"">
";

    internal const string ContentTypesDefault = @"<Default Extension=""rels"" ContentType=""application/vnd.openxmlformats-package.relationships+xml"" />
    <Default Extension=""xml"" ContentType=""application/xml"" />
";

    internal const string ContentTypesEnd = "</Types>";
}