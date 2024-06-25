using System.Diagnostics.CodeAnalysis;

namespace FastExcelSlim;

public class OpenXmlFormatterException : Exception
{
    public OpenXmlFormatterException(string message)
        : base(message)
    {
    }

    public OpenXmlFormatterException(string message, Exception innerException)
        : base(message, innerException)
    {
    }

    [DoesNotReturn]
    public static void ThrowMessage(string message)
    {
        throw new OpenXmlFormatterException(message);
    }

    [DoesNotReturn]
    public static void ThrowRegisterInProviderFailed(Type type, Exception innerException)
    {
        throw new OpenXmlFormatterException($"{type.FullName} is failed in provider at creating formatter.", innerException);
    }

    [DoesNotReturn]
    public static void ThrowNotRegisteredInProvider(Type type)
    {
        throw new OpenXmlFormatterException($"{type.FullName} is not registered in this provider.");
    }

}