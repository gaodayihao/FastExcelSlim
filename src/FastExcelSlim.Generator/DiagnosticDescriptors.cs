using Microsoft.CodeAnalysis;

namespace FastExcelSlim.Generator;

internal static class DiagnosticDescriptors
{
    const string Category = "GenerateFastExcelSlim";

    public static readonly DiagnosticDescriptor MustBePartial = new(
        id: "FES001",
        title: "OpenXmlSerializable object must be partial",
        messageFormat: "The OpenXmlSerializable object '{0}' must be partial",
        category: Category,
        defaultSeverity: DiagnosticSeverity.Error,
        isEnabledByDefault: true);
}