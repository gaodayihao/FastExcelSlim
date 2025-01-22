#if NETSTANDARD2_0
namespace FastExcelSlim;

[AttributeUsage(AttributeTargets.Method, Inherited = false)]
internal sealed class DoesNotReturnAttribute : Attribute
{
}
#endif