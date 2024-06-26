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

        foreach (var member in Members)
        {
            if (!_reference.KnownTypes.Contains(member.MemberType))
            {
                context.ReportDiagnostic(Diagnostic.Create(DiagnosticDescriptors.MemberCantSerializeType, member.GetLocation(syntax), Symbol.Name, member.Name, member.MemberType.FullyQualifiedToString()));
                noError = false;
            }
        }

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

        var scopedRef = (context.IsCSharp11OrGreater())
            ? "scoped ref"
            : "ref";

        const string constraint = "where TBufferWriter : System.Buffers.IBufferWriter<byte>";

        writer.AppendLine($$"""
partial {{classOrStructOrRecord}} {{TypeName}} : IOpenXmlWritable<{{TypeName}}>
{
    public static int ColumnCount => {{Members.Length}};

    public static string? SheetName => {{GetSheetName()}};

    static partial void StaticConstructor();

    static {{Symbol.Name}}()
    {
        global::FastExcelSlim.OpenXmlFormatterProvider.Register<{{TypeName}}>();
        StaticConstructor();
    }

    public static void RegisterFormatter()
    {
        if (!global::FastExcelSlim.OpenXmlFormatterProvider.IsRegistered<{{TypeName}}>())
        {
            global::FastExcelSlim.OpenXmlFormatterProvider.Register(new global::FastExcelSlim.OpenXmlFormatter<{{TypeName}}>());
        }
    }

    public static void WriteCell<TBufferWriter>(
        {{scopedRef}} global::Utf8StringInterpolation.Utf8StringWriter<TBufferWriter> writer,
        global::FastExcelSlim.OpenXmlStyles styles,
        int rowIndex,
        {{scopedRef}} {{TypeName}} value) {{constraint}}
    {
{{EmitWriteCell("        ").NewLine()}}
    }
    
    public static void WriteColumns<TBufferWriter>(
        {{scopedRef}} global::Utf8StringInterpolation.Utf8StringWriter<TBufferWriter> writer) {{constraint}}
    {
{{EmitWriteColumns("        ").NewLine()}}
    }
    
    public static void WriteHeaders<TBufferWriter>
        ({{scopedRef}} global::Utf8StringInterpolation.Utf8StringWriter<TBufferWriter> writer,
        global::FastExcelSlim.OpenXmlStyles styles) {{constraint}}
    {
{{EmitWriteHeaders("        ").NewLine()}}
    }
}
""");
    }

    private string GetSheetName()
    {
        const string nullSheetName = "null";

        var attr = Symbol.GetAttribute(_reference.OpenXmlWritableAttribute)!;
        var args = attr.ConstructorArguments;
        if (args.Length == 1)
        {
            if (args[0].Value is string sheetName)
            {
                return $"\"{sheetName}\"";
            }
        }

        return nullSheetName;
    }

    private IEnumerable<string> EmitWriteCell(string indent)
    {
        for (int i = 0; i < Members.Length; i++)
        {
            yield return $"{indent}writer.WriteCell(styles, value.{Members[i].Name}, rowIndex, {i + 1}, nameof({Members[i].Name}), ref value);";
        }
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

    public override string ToString()
    {
        return TypeName;
    }
}