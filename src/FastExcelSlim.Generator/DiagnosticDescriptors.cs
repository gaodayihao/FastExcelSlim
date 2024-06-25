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

    public static readonly DiagnosticDescriptor MembersCountOver16384 = new(
        id: "FES002",
        title: "Members count limit",
        messageFormat: "The OpenXmlWritable object '{0}' member count is '{1}', however limit size is 16384",
        category: Category,
        defaultSeverity: DiagnosticSeverity.Error,
        isEnabledByDefault: true);

    public static readonly DiagnosticDescriptor AllMembersMustAnnotateOrder = new(
        id: "FES003",
        title: "All members must annotate OpenXmlOrder",
        messageFormat: "The OpenXmlWritable object '{0}' member '{1}' is not annotated OpenXmlOrder",
        category: Category,
        defaultSeverity: DiagnosticSeverity.Error,
        isEnabledByDefault: true);

    public static readonly DiagnosticDescriptor DuplicateOrderDoesNotAllow = new(
        id: "FES004",
        title: "All members order must be unique",
        messageFormat: "The OpenXmlWritable object '{0}' member '{1}' is duplicated order between '{2}'.",
        category: Category,
        defaultSeverity: DiagnosticSeverity.Error,
        isEnabledByDefault: true);

    public static readonly DiagnosticDescriptor AllMembersMustBeContinuousNumber = new(
        id: "FES005",
        title: "All OpenXmlWritable members must be continuous number from zero",
        messageFormat: "The OpenXmlWritable object '{0}' member '{1}' is not continuous number from zero",
        category: Category,
        defaultSeverity: DiagnosticSeverity.Error,
        isEnabledByDefault: true);


    public static readonly DiagnosticDescriptor MemberCantSerializeType = new(
        id: "FES006",
        title: "Member can't serialize type",
        messageFormat: "The OpenXmlWritable object '{0}' member '{1}' type is '{2}' that can't serialize",
        category: Category,
        defaultSeverity: DiagnosticSeverity.Error,
        isEnabledByDefault: true);
}