using Microsoft.CodeAnalysis.CSharp.Syntax;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;

namespace FastExcelSlim.Generator;

public partial class FastExcelSlimGenerator
{
    static void Generate(TypeDeclarationSyntax syntax, Compilation compilation, string? serializationInfoLogDirectoryPath, IGeneratorContext context)
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


    }

    static bool IsPartial(TypeDeclarationSyntax typeDeclaration)
    {
        return typeDeclaration.Modifiers.Any(m => m.IsKind(SyntaxKind.PartialKeyword));
    }
}