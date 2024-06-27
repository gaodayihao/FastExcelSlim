using System.Text;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp.Syntax;

namespace FastExcelSlim.Generator;

internal class TypeMeta
{
    private readonly ReferenceSymbols _reference;

    public MemberMeta[] Members { get; }

    public bool IsRecord { get; }

    public bool IsValueType { get; }

    public TypeMeta(ReferenceSymbols reference, INamedTypeSymbol symbol)
    {
        _reference = reference;
        Symbol = symbol;

        TypeName = symbol.ToDisplayString(SymbolDisplayFormat.MinimallyQualifiedFormat);

        Members = symbol.GetAllMembers()
            .Where(x => x is (IFieldSymbol or IPropertySymbol) and { IsStatic: false, IsImplicitlyDeclared: false, CanBeReferencedByName: true })
            .Reverse()
            .DistinctBy(x => x.Name)
            .Reverse()
            .Where(x =>
            {
                var include = x.ContainsAttribute(reference.OpenXmlPropertyAttribute);
                var ignore = x.ContainsAttribute(reference.OpenXmlIgnoreAttribute);
                if (ignore) return false;
                if (include) return true;
                return x.DeclaredAccessibility is Accessibility.Public;
            })
            .Where(x =>
            {
                if (x is IPropertySymbol p)
                {
                    if (p.GetMethod == null && p.SetMethod != null)
                    {
                        return false;
                    }

                    if (p.IsIndexer) return false;
                }

                return true;
            })
            .Select((x, i) => new MemberMeta(x, reference, i))
            .OrderBy(x => x.Order)
            .ToArray();

        IsRecord = symbol.IsRecord;
        IsValueType = symbol.IsValueType;
    }

    public string TypeName { get; }

    public INamedTypeSymbol Symbol { get; }

    public bool Validate(TypeDeclarationSyntax syntax, IGeneratorContext context)
    {
        var noError = true;
        if (Members.Length > 16_384)
        {
            context.ReportDiagnostic(Diagnostic.Create(DiagnosticDescriptors.MembersCountOver16384, syntax.Identifier.GetLocation(), Symbol.Name, Members.Length));
            noError = false;
        }

        foreach (var member in Members)
        {
            if (member.MemberType.TypeKind != TypeKind.Enum && !_reference.KnownTypes.Contains(member.MemberType))
            {
                context.ReportDiagnostic(Diagnostic.Create(DiagnosticDescriptors.MemberCantSerializeType, member.GetLocation(syntax), Symbol.Name, member.Name, member.MemberType.FullyQualifiedToString()));
                noError = false;
            }

            if (member.ContainsAttribute(_reference.OpenXmlEnumFormatAttribute) && member.MemberType.TypeKind != TypeKind.Enum)
            {
                context.ReportDiagnostic(Diagnostic.Create(DiagnosticDescriptors.MemberNotAnEnumType, member.GetLocation(syntax), Symbol.Name, member.Name));
                noError = false;
            }
        }

        //order
        if (Members.Any(m => m.HasExplicitOrder))
        {
            if (!Members.All(m => m.HasExplicitOrder))
            {
                foreach (var item in Members.Where(m => !m.HasExplicitOrder))
                {
                    context.ReportDiagnostic(Diagnostic.Create(DiagnosticDescriptors.AllMembersMustAnnotateOrder, item.GetLocation(syntax), Symbol.Name, item.Name));
                }
                noError = false;
                goto RETURN;
            }
        }

        var orderSet = new Dictionary<int, MemberMeta>(Members.Length);
        foreach (var member in Members)
        {
            if (orderSet.TryGetValue(member.Order, out var duplicateMember))
            {
                context.ReportDiagnostic(Diagnostic.Create(DiagnosticDescriptors.DuplicateOrderDoesNotAllow, member.GetLocation(syntax), Symbol.Name, member.Name, duplicateMember.Name));
                noError = false;
            }
            else
            {
                orderSet.Add(member.Order, member);
            }
        }

        if (!noError) goto RETURN;

        var expectedOrder = 0;
        foreach (var member in Members)
        {
            if (member.Order != expectedOrder)
            {
                context.ReportDiagnostic(Diagnostic.Create(DiagnosticDescriptors.AllMembersMustBeContinuousNumber, member.GetLocation(syntax), Symbol.Name, member.Name));
                noError = false;
                break;
            }
            expectedOrder++;
        }

    RETURN:
        return noError;
    }

    public void Emit(StringBuilder writer, IGeneratorContext context)
    {
        var classOrStructOrRecord = (IsRecord, IsValueType) switch
        {
            (true, true) => "record struct",
            (true, false) => "record",
            (false, true) => "struct",
            (false, false) => "class",
        };

        string staticReturnVoidMethod, staticReturnIntMethod, staticReturnStringMethod, staticReturnOptionMethod, registerBody, registerT, constraint;
        var scopedRef = (context.IsCSharp11OrGreater())
            ? "scoped ref"
            : "ref";

        if (!context.IsNet7OrGreater)
        {
            staticReturnVoidMethod = "internal static void ";
            staticReturnIntMethod = "internal static int ";
            staticReturnStringMethod = "internal static string? ";
            staticReturnOptionMethod = "internal static global::FastExcelSlim.OpenXml.OpenXmlExcelOptions ";
            registerBody = $"global::FastExcelSlim.OpenXmlFormatterProvider.Register(new {Symbol.Name}Formatter());";
            registerT = "RegisterFormatter();";
            constraint = " where TBufferWriter : System.Buffers.IBufferWriter<byte>";
        }
        else
        {
            staticReturnVoidMethod = $"static void IOpenXmlWritable<{TypeName}>.";
            staticReturnIntMethod = $"static int IOpenXmlWritable<{TypeName}>.";
            staticReturnStringMethod = $"static string? IOpenXmlWritable<{TypeName}>.";
            staticReturnOptionMethod = $"static global::FastExcelSlim.OpenXml.OpenXmlExcelOptions IOpenXmlWritable<{TypeName}>.";
            registerBody = $"global::FastExcelSlim.OpenXmlFormatterProvider.Register(new global::FastExcelSlim.OpenXmlFormatter<{TypeName}>());";
            registerT = $"global::FastExcelSlim.OpenXmlFormatterProvider.Register<{TypeName}>();";
            constraint = string.Empty;
        }

        writer.AppendLine($$"""
partial {{classOrStructOrRecord}} {{TypeName}} : IOpenXmlWritable<{{TypeName}}>
{
    [global::FastExcelSlim.Internal.Preserve]
    {{staticReturnIntMethod}}ColumnCount => {{Members.Length}};

    [global::FastExcelSlim.Internal.Preserve]
    {{staticReturnStringMethod}}SheetName => {{GetSheetName()}};

    static partial void StaticConstructor();

    static {{Symbol.Name}}()
    {
        {{registerT}}
        StaticConstructor();
    }
    
    [global::FastExcelSlim.Internal.Preserve]
    {{staticReturnVoidMethod}}RegisterFormatter()
    {
        if (!global::FastExcelSlim.OpenXmlFormatterProvider.IsRegistered<{{TypeName}}>())
        {
            {{registerBody}}
        }
    }

    [global::FastExcelSlim.Internal.Preserve]
    {{staticReturnVoidMethod}}WriteCell<TBufferWriter>(
        {{scopedRef}} global::Utf8StringInterpolation.Utf8StringWriter<TBufferWriter> writer,
        global::FastExcelSlim.OpenXmlStyles styles,
        int rowIndex,
        {{scopedRef}} {{TypeName}} value){{constraint}}
    {
{{EmitWriteCell("        ").NewLine()}}
    }
    
    [global::FastExcelSlim.Internal.Preserve]
    {{staticReturnVoidMethod}}WriteColumns<TBufferWriter>(
        {{scopedRef}} global::Utf8StringInterpolation.Utf8StringWriter<TBufferWriter> writer){{constraint}}
    {
{{EmitWriteColumns("        ").NewLine()}}
    }
    
    [global::FastExcelSlim.Internal.Preserve]
    {{staticReturnVoidMethod}}WriteHeaders<TBufferWriter>
        ({{scopedRef}} global::Utf8StringInterpolation.Utf8StringWriter<TBufferWriter> writer,
        global::FastExcelSlim.OpenXmlStyles styles){{constraint}}
    {
{{EmitWriteHeaders("        ").NewLine()}}
    }
    
    [global::FastExcelSlim.Internal.Preserve]
    {{staticReturnOptionMethod}}GetOptions()
    {
        {{EmitWriteOptions()}}
    }
}
""");

        if (!context.IsNet7OrGreater)
        {
            // add formatter(can not use MemoryPackableFormatter)

            var code = $$"""

partial {{classOrStructOrRecord}} {{TypeName}}
{
    sealed class {{Symbol.Name}}Formatter : IOpenXmlFormatter<{{TypeName}}>
    {
        [global::FastExcelSlim.Internal.Preserve]
        void IOpenXmlFormatter<{{TypeName}}>.WriteCell<TBufferWriter>({{scopedRef}} global::Utf8StringInterpolation.Utf8StringWriter<TBufferWriter> writer, OpenXmlStyles styles, int rowIndex,
            {{scopedRef}} {{TypeName}} value)
        {
            WriteCell(ref writer, styles, rowIndex, ref value);
        }
    
        [global::FastExcelSlim.Internal.Preserve]
        void IOpenXmlFormatter<{{TypeName}}>.WriteColumns<TBufferWriter>({{scopedRef}} global::Utf8StringInterpolation.Utf8StringWriter<TBufferWriter> writer)
        {
            WriteColumns(ref writer);
        }
    
        [global::FastExcelSlim.Internal.Preserve]
        void IOpenXmlFormatter<{{TypeName}}>.WriteHeaders<TBufferWriter>({{scopedRef}} global::Utf8StringInterpolation.Utf8StringWriter<TBufferWriter> writer, OpenXmlStyles styles)
        {
            WriteHeaders(ref writer, styles);
        }
    
        [global::FastExcelSlim.Internal.Preserve]
        int IOpenXmlFormatter<{{TypeName}}>.ColumnCount => ColumnCount;
    
        [global::FastExcelSlim.Internal.Preserve]
        string? IOpenXmlFormatter<{{TypeName}}>.SheetName => SheetName;
        
        [global::FastExcelSlim.Internal.Preserve]
        global::FastExcelSlim.OpenXml.OpenXmlExcelOptions IOpenXmlFormatter<{{TypeName}}>.GetOptions() => GetOptions();
    }
}
""";
            writer.AppendLine(code);
        }
    }

    private string GetSheetName()
    {
        const string nullSheetName = "null";
        var sheetName = GetAttributeNamedArgumentValue<string?>(_reference.OpenXmlWritableAttribute, "SheetName", null);
        if (!string.IsNullOrEmpty(sheetName))
        {
            return $"\"{sheetName}\"";
        }

        return nullSheetName;
    }

    private IEnumerable<string> EmitWriteCell(string indent)
    {
        for (int i = 0; i < Members.Length; i++)
        {
            var member = Members[i];
            if (member.MemberType.TypeKind == TypeKind.Enum &&
                member.ContainsAttribute(_reference.OpenXmlEnumFormatAttribute))
            {
                var formatAttr = member.GetAttribute(_reference.OpenXmlEnumFormatAttribute);
                var format = (string)formatAttr!.ConstructorArguments[0].Value!;
                yield return $"{indent}writer.WriteCell(styles, value.{member.Name}, rowIndex, {i + 1}, nameof({member.Name}), ref value, \"{format}\");";
            }
            else
            {
                yield return $"{indent}writer.WriteCell(styles, value.{member.Name}, rowIndex, {i + 1}, nameof({member.Name}), ref value);";
            }
        }
    }

    private string EmitWriteOptions()
    {
        var autoFilter = GetAttributeNamedArgumentValue(_reference.OpenXmlWritableAttribute, "AutoFilter", false);
        var freezeHeader = GetAttributeNamedArgumentValue(_reference.OpenXmlWritableAttribute, "FreezeHeader", false);

        return $"return new global::FastExcelSlim.OpenXml.OpenXmlExcelOptions({freezeHeader.ToString().ToLower()}, {autoFilter.ToString().ToLower()});";
    }

    private IEnumerable<string> EmitWriteColumns(string indent)
    {
        for (int i = 0; i < Members.Length; i++)
        {
            yield return $"{indent}writer.WriteColumn({i + 1}, {Members[i].GetColumnWidth()});";
        }
    }

    private IEnumerable<string> EmitWriteHeaders(string indent)
    {
        for (int i = 0; i < Members.Length; i++)
        {
            yield return $"{indent}writer.WriteHeader(styles, \"{Members[i].GetColumnName()}\", nameof({Members[i].Name}), {i + 1});";
        }
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

    public override string ToString()
    {
        return TypeName;
    }
}