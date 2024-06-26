using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp.Syntax;

namespace FastExcelSlim.Generator;

internal class MemberMeta
{
    internal int DefaultColumnWidth = 15;

    private readonly ReferenceSymbols _references;

    public MemberMeta(ISymbol symbol, ReferenceSymbols references, int sequentialOrder)
    {
        _references = references;
        Symbol = symbol;
        Name = symbol.Name;
        Order = sequentialOrder;
        var orderAttr = symbol.GetAttribute(references.OpenXmlOrderAttribute);
        if (orderAttr != null)
        {
            Order = (int)(orderAttr.ConstructorArguments[0].Value ?? sequentialOrder);
            HasExplicitOrder = true;
        }

        if (symbol is IFieldSymbol f)
        {
            IsProperty = false;
            IsField = true;
            MemberType = f.Type;
        }
        else if (symbol is IPropertySymbol p)
        {
            IsProperty = true;
            IsField = false;
            MemberType = p.Type;
        }
        else
        {
            throw new Exception("member is not field or property.");
        }
    }

    public ISymbol Symbol { get; }

    public bool IsField { get; }

    public bool IsProperty { get; }

    public int Order { get; }

    public bool HasExplicitOrder { get; }

    public string Name { get; }

    public ITypeSymbol MemberType { get; }

    public Location GetLocation(TypeDeclarationSyntax fallback)
    {
        var location = Symbol.Locations.FirstOrDefault() ?? fallback.Identifier.GetLocation();
        return location;
    }

    public string GetColumnName()
    {
        return GetAttributeNamedArgumentValue<string>(_references.OpenXmlPropertyAttribute, "ColumnName", Name);
    }

    public int GetColumnWidth()
    {
        return GetAttributeNamedArgumentValue<int>(_references.OpenXmlPropertyAttribute, "Width", DefaultColumnWidth);
    }

    private T GetAttributeNamedArgumentValue<T>(INamedTypeSymbol attributeSymbol, string argumentName, T defaultValue)
    {
        var propertyAttr = Symbol.GetAttribute(attributeSymbol);

        if (propertyAttr is { NamedArguments.IsEmpty: false })
        {
            foreach (var kvp in propertyAttr.NamedArguments)
            {
                var argName = kvp.Key;
                var typedConstant = kvp.Value;
                if (argName == argumentName && typedConstant.Value is T value)
                {
                    return value;
                }
            }
        }

        return defaultValue;
    }

    public AttributeData? GetAttribute(INamedTypeSymbol attribute)
    {
        return Symbol.GetAttribute(attribute);
    }

    public bool ContainsAttribute(INamedTypeSymbol attribute)
    {
        return Symbol.ContainsAttribute(attribute);
    }
}