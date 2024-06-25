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
        var propertyAttr = Symbol.GetAttribute(_references.OpenXmlPropertyAttribute);
        if (propertyAttr is { ConstructorArguments.Length: > 0 })
        {
            if (propertyAttr.ConstructorArguments[0].Value is string columnName)
            {
                return columnName;
            }
        }

        return Name;
    }

    public int GetColumnWidth()
    {
        var propertyAttr = Symbol.GetAttribute(_references.OpenXmlPropertyAttribute);

        if (propertyAttr != null)
        {
            switch (propertyAttr.ConstructorArguments.Length)
            {
                case 2:
                    {
                        if (propertyAttr.ConstructorArguments[1].Value is int width)
                        {
                            return width;
                        }
                        break;
                    }
                case 1:
                    {
                        if (propertyAttr.ConstructorArguments[0].Value is int width)
                        {
                            return width;
                        }
                        break;
                    }
            }
        }

        return DefaultColumnWidth;
    }
}