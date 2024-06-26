using Microsoft.CodeAnalysis;

namespace FastExcelSlim.Generator;

internal class ReferenceSymbols
{
    public ReferenceSymbols(Compilation compilation)
    {
        Compilation = compilation;

        OpenXmlWritableAttribute = GetTypeByMetadataName(FastExcelSlimGenerator.OpenXmlWritableAttributeFullName);
        OpenXmlPropertyAttribute = GetTypeByMetadataName("FastExcelSlim.OpenXmlPropertyAttribute");
        OpenXmlIgnoreAttribute = GetTypeByMetadataName("FastExcelSlim.OpenXmlIgnoreAttribute");
        OpenXmlOrderAttribute = GetTypeByMetadataName("FastExcelSlim.OpenXmlOrderAttribute");
        OpenXmlEnumFormatAttribute = GetTypeByMetadataName("FastExcelSlim.OpenXmlEnumFormatAttribute");

        KnownTypes = new WellKnownTypes(this);
    }

    public Compilation Compilation { get; }

    public INamedTypeSymbol OpenXmlWritableAttribute { get; }
    public INamedTypeSymbol OpenXmlPropertyAttribute { get; }
    public INamedTypeSymbol OpenXmlIgnoreAttribute { get; }
    public INamedTypeSymbol OpenXmlOrderAttribute { get; }
    public INamedTypeSymbol OpenXmlEnumFormatAttribute { get; }

    public WellKnownTypes KnownTypes { get; }

    internal INamedTypeSymbol GetTypeByMetadataName(string metadataName)
    {
        var symbol = Compilation.GetTypeByMetadataName(metadataName);
        if (symbol == null)
        {
            throw new InvalidOperationException($"Type {metadataName} is not found in compilation.");
        }
        return symbol;
    }
}