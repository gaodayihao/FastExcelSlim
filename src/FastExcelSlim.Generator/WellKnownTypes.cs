using Microsoft.CodeAnalysis;
// ReSharper disable InconsistentNaming

namespace FastExcelSlim.Generator;

internal class WellKnownTypes
{
    private readonly ReferenceSymbols _parent;
    private readonly HashSet<ITypeSymbol> _knownTypes;

    public WellKnownTypes(ReferenceSymbols parent, IGeneratorContext context)
    {
        _parent = parent;
        System_Nullable = GetTypeByMetadataName("System.Nullable`1");
        System_Boolean = GetTypeByMetadataName("System.Boolean");
        System_Guid = GetTypeByMetadataName("System.Guid");
        if (context.IsNet7OrGreater)
        {
            System_TimeOnly = GetTypeByMetadataName("System.TimeOnly");
            System_DateOnly = GetTypeByMetadataName("System.DateOnly");
        }
        System_DateTime = GetTypeByMetadataName("System.DateTime");
        System_Byte = GetTypeByMetadataName("System.Byte");
        System_Decimal = GetTypeByMetadataName("System.Decimal");
        System_Double = GetTypeByMetadataName("System.Double");
        System_Int16 = GetTypeByMetadataName("System.Int16");
        System_Int32 = GetTypeByMetadataName("System.Int32");
        System_Int64 = GetTypeByMetadataName("System.Int64");
        System_SByte = GetTypeByMetadataName("System.SByte");
        System_Single = GetTypeByMetadataName("System.Single");
        System_UInt16 = GetTypeByMetadataName("System.UInt16");
        System_UInt32 = GetTypeByMetadataName("System.UInt32");
        System_UInt64 = GetTypeByMetadataName("System.UInt64");
        System_String = GetTypeByMetadataName("System.String");

        _knownTypes = new HashSet<ITypeSymbol>(new[]
        {
            System_Boolean,
            System_Guid,
            System_DateTime,
            System_Byte,
            System_Decimal,
            System_Double,
            System_Int16,
            System_Int32,
            System_Int64,
            System_SByte,
            System_Single,
            System_UInt16,
            System_UInt32,
            System_UInt64,
            System_String
        }, SymbolEqualityComparer.Default);
        if (context.IsNet7OrGreater)
        {
            _knownTypes.Add(System_TimeOnly!);
            _knownTypes.Add(System_DateOnly!);
        }
    }

    //System.Nullable
    public INamedTypeSymbol System_Nullable { get; }
    public INamedTypeSymbol System_Boolean { get; }
    public INamedTypeSymbol System_Guid { get; }
    public INamedTypeSymbol? System_TimeOnly { get; }
    public INamedTypeSymbol? System_DateOnly { get; }
    public INamedTypeSymbol System_DateTime { get; }
    public INamedTypeSymbol System_Byte { get; }
    public INamedTypeSymbol System_Decimal { get; }
    public INamedTypeSymbol System_Double { get; }
    public INamedTypeSymbol System_Int16 { get; }
    public INamedTypeSymbol System_Int32 { get; }
    public INamedTypeSymbol System_Int64 { get; }
    public INamedTypeSymbol System_SByte { get; }
    public INamedTypeSymbol System_Single { get; }
    public INamedTypeSymbol System_UInt16 { get; }
    public INamedTypeSymbol System_UInt32 { get; }
    public INamedTypeSymbol System_UInt64 { get; }
    public INamedTypeSymbol System_String { get; }

    public bool Contains(ITypeSymbol symbol)
    {
        if (symbol is INamedTypeSymbol { IsGenericType: true } nts && SymbolEqualityComparer.Default.Equals(nts.OriginalDefinition, System_Nullable))
        {
            symbol = nts.TypeArguments[0];
        }

        var contains = _knownTypes.Contains(symbol);
        return contains;
    }

    internal INamedTypeSymbol GetTypeByMetadataName(string metadataName) => _parent.GetTypeByMetadataName(metadataName);
}