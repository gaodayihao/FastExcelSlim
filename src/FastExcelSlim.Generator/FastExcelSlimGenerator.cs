using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis.CSharp.Syntax;

namespace FastExcelSlim.Generator;

// dotnet/runtime generators.

// https://github.com/dotnet/runtime/blob/main/src/libraries/System.Text.RegularExpressions/gen/
// https://github.com/dotnet/runtime/tree/main/src/libraries/System.Text.Json/gen
// https://github.com/dotnet/runtime/tree/main/src/libraries/System.Private.CoreLib/gen
// https://github.com/dotnet/runtime/tree/main/src/libraries/Microsoft.Extensions.Logging.Abstractions/gen
// https://github.com/dotnet/runtime/tree/main/src/libraries/System.Runtime.InteropServices.JavaScript/gen/JSImportGenerator
// https://github.com/dotnet/runtime/tree/main/src/libraries/System.Runtime.InteropServices/gen/LibraryImportGenerator
// https://github.com/dotnet/runtime/tree/main/src/tests/Common/XUnitWrapperGenerator

// documents, blogs.

// https://github.com/dotnet/roslyn/blob/main/docs/features/incremental-generators.md
// https://andrewlock.net/creating-a-source-generator-part-1-creating-an-incremental-source-generator/
// https://qiita.com/WiZLite/items/48f37278cf13be899e40
// https://zenn.dev/pcysl5edgo/articles/6d9be0dd99c008
// https://neue.cc/2021/05/08_600.html
// https://www.thinktecture.com/en/net/roslyn-source-generators-introduction/

// for check generated file
// <EmitCompilerGeneratedFiles>true</EmitCompilerGeneratedFiles>
// <CompilerGeneratedFilesOutputPath>Generated</CompilerGeneratedFilesOutputPath>

[Generator(LanguageNames.CSharp)]
public partial class FastExcelSlimGenerator : IIncrementalGenerator
{
    public const string OpenXmlWritableAttributeFullName = "FastExcelSlim.OpenXmlWritableAttribute";

    public void Initialize(IncrementalGeneratorInitializationContext context)
    {
        RegisterFastExcelSlim(context);
    }

    private void RegisterFastExcelSlim(IncrementalGeneratorInitializationContext context)
    {
        var typeDeclarations = context.SyntaxProvider.ForAttributeWithMetadataName(
            OpenXmlWritableAttributeFullName,
            predicate: static (node, _) => node is ClassDeclarationSyntax
                or StructDeclarationSyntax
                or RecordDeclarationSyntax,
            transform: static (context, _) => (TypeDeclarationSyntax)context.TargetNode);

        var parseOptions = context.ParseOptionsProvider.Select((parseOptions, _) =>
        {
            var csOptions = (CSharpParseOptions)parseOptions;
            var langVersion = csOptions.LanguageVersion;
            var net7 = csOptions.PreprocessorSymbolNames.Contains("NET7_0_OR_GREATER");
            return (lanVersion: langVersion, net7);
        });

        var source = typeDeclarations
            .Combine(context.CompilationProvider)
            .WithComparer(Comparer.Instance)
            .Combine(parseOptions);

        context.RegisterSourceOutput(source, static (context, source) =>
        {
            var (typeDeclaration, compilation) = source.Left;
            var (langVersion, net7) = source.Right;

            Generate(typeDeclaration, compilation, new GeneratorContext(context, langVersion, net7));
        });
    }

    private class Comparer : IEqualityComparer<(TypeDeclarationSyntax, Compilation)>
    {
        public static readonly Comparer Instance = new();

        public bool Equals((TypeDeclarationSyntax, Compilation) x, (TypeDeclarationSyntax, Compilation) y)
        {
            return x.Item1.Equals(y.Item1);
        }

        public int GetHashCode((TypeDeclarationSyntax, Compilation) obj)
        {
            return obj.Item1.GetHashCode();
        }
    }

    class GeneratorContext : IGeneratorContext
    {
        private readonly SourceProductionContext _context;

        public GeneratorContext(SourceProductionContext context, LanguageVersion languageVersion, bool isNet70OrGreater)
        {
            _context = context;
            LanguageVersion = languageVersion;
            IsNet7OrGreater = isNet70OrGreater;
        }

        public CancellationToken CancellationToken => _context.CancellationToken;

        public LanguageVersion LanguageVersion { get; }

        public bool IsNet7OrGreater { get; }

        public void AddSource(string hintName, string source)
        {
            _context.AddSource(hintName, source);
        }

        public void ReportDiagnostic(Diagnostic diagnostic)
        {
            _context.ReportDiagnostic(diagnostic);
        }
    }
}