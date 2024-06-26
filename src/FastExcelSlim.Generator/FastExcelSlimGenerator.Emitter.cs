﻿using Microsoft.CodeAnalysis.CSharp.Syntax;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using System.Text;

namespace FastExcelSlim.Generator;

public partial class FastExcelSlimGenerator
{
    static void Generate(TypeDeclarationSyntax syntax, Compilation compilation, IGeneratorContext context)
    {
        var semanticModel = compilation.GetSemanticModel(syntax.SyntaxTree);

        var typeSymbol = semanticModel.GetDeclaredSymbol(syntax, context.CancellationToken);
        if (typeSymbol == null)
        {
            return;
        }

        // verify is partial
        if (!IsPartial(syntax))
        {
            context.ReportDiagnostic(Diagnostic.Create(DiagnosticDescriptors.MustBePartial, syntax.Identifier.GetLocation(), typeSymbol.Name));
            return;
        }

        var reference = new ReferenceSymbols(compilation);

        var typeMeta = new TypeMeta(reference, typeSymbol);

        if (!typeMeta.Validate(syntax, context))
        {
            return;
        }

        var fullType = typeSymbol.ToDisplayString(SymbolDisplayFormat.FullyQualifiedFormat)
            .Replace("global::", "")
            .Replace("<", "_")
            .Replace(">", "_");

        var sb = new StringBuilder();

        sb.AppendLine(
@"// <auto-generated/>
#nullable enable
#pragma warning disable CS0108 // hides inherited member
#pragma warning disable CS0162 // Unreachable code
#pragma warning disable CS0164 // This label has not been referenced
#pragma warning disable CS0219 // Variable assigned but never used
#pragma warning disable CS8600 // Converting null literal or possible null value to non-nullable type.
#pragma warning disable CS8601 // Possible null reference assignment
#pragma warning disable CS8602 // Dereference of a possibly null reference.
#pragma warning disable CS8604 // Possible null reference argument for parameter
#pragma warning disable CS8619 // Nullability of reference types in value doesn't match target type.
#pragma warning disable CS8620 // Argument cannot be used for parameter due to differences in the nullability of reference types.
#pragma warning disable CS8631 // The type cannot be used as type parameter in the generic type or method
#pragma warning disable CS8765 // Nullability of type of parameter
#pragma warning disable CS9074 // The 'scoped' modifier of parameter doesn't match overridden or implemented member
#pragma warning disable CA1050 // Declare types in namespaces.

using System;
using System.Buffers;
using FastExcelSlim;
using FastExcelSlim.Extensions;
using Utf8StringInterpolation;
");

        var ns = typeMeta.Symbol.ContainingNamespace;
        if (!ns.IsGlobalNamespace)
        {
            if (context.IsCSharp10OrGreater())
            {
                sb.AppendLine($"namespace {ns};");
            }
            else
            {
                sb.AppendLine($"namespace {ns} {{");
            }
        }
        sb.AppendLine();

        BuildDebugInfo(sb, typeMeta);

        typeMeta.Emit(sb, context);

        if (!ns.IsGlobalNamespace && !context.IsCSharp10OrGreater())
        {
            sb.AppendLine("}");
        }

        var code = sb.ToString();
        context.AddSource($"{fullType}.OpenXmlFormatter.g.cs", code);
    }

    static void BuildDebugInfo(StringBuilder sb, TypeMeta type)
    {
        string WithEscape(ISymbol symbol)
        {
            var str = symbol.FullyQualifiedToString().Replace("global::", "");
            return str.Replace("<", "&lt;").Replace(">", "&gt;");
        }

        sb.AppendLine("/// <remarks>");
        sb.AppendLine("/// <code>");

        foreach (var item in type.Members)
        {
            sb.Append("/// <b>");
            sb.Append(WithEscape(item.MemberType));
            sb.Append("</b>");

            sb.Append(" ");
            sb.Append(item.Name);
            sb.AppendLine("<br/>");
        }
        sb.AppendLine("/// </code>");
        sb.AppendLine("/// </remarks>");
    }

    static bool IsPartial(TypeDeclarationSyntax typeDeclaration)
    {
        return typeDeclaration.Modifiers.Any(m => m.IsKind(SyntaxKind.PartialKeyword));
    }
}