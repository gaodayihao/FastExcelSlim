﻿using System.Buffers;
using System.Diagnostics.CodeAnalysis;
using System.Reflection;
using Utf8StringInterpolation;

namespace FastExcelSlim;

public static class OpenXmlFormatterProvider
{
    public static bool IsRegistered<T>() => Check<T>.Registered;

#if NET7_0_OR_GREATER
    public static void Register<T>() where T : IOpenXmlFormatterRegister
    {
        T.RegisterFormatter();
    }
#endif

    public static void Register<T>(IOpenXmlFormatter<T> formatter)
    {
        Check<T>.Registered = true;
        Cache<T>.Formatter = formatter;
    }

    internal static IOpenXmlFormatter<T> GetFormatter<T>()
    {
        return Cache<T>.Formatter ?? new ErrorOpenXmlFormatter<T>();
    }

    static class Check<T>
    {
        public static bool Registered;
    }

    static class Cache<T>
    {
        public static IOpenXmlFormatter<T>? Formatter;

        static Cache()
        {
            if (Check<T>.Registered) return;

            try
            {
                if (TryInvokeRegisterFormatter(typeof(T)))
                {
                    return;
                }
            }
            catch (Exception e)
            {
                Formatter = new ErrorOpenXmlFormatter<T>(e);
                return;
            }

            Formatter = new ErrorOpenXmlFormatter<T>();
            Check<T>.Registered = true;
        }
    }

    static bool TryInvokeRegisterFormatter(Type type)
    {
        if (typeof(IOpenXmlFormatterRegister).IsAssignableFrom(type))
        {
            var m = type.GetMethod("FastExcelSlim.IOpenXmlFormatterRegister.RegisterFormatter", BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Static);
            if (m == null)
            {
                // Roslyn3.11 generator generate public static method
                m = type.GetMethod("RegisterFormatter", BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Static);
            }
            if (m == null)
            {
                throw new InvalidOperationException("Type implements IOpenXmlWritable but can not found RegisterFormatter. Type: " + type.FullName);
            }

            m.Invoke(null, null);
            return true;
        }

        return false;
    }
}

internal sealed class ErrorOpenXmlFormatter<T> : IOpenXmlFormatter<T>
{
    private readonly Exception? _exception;
    private readonly string? _message;

    public ErrorOpenXmlFormatter()
    {
    }

    public ErrorOpenXmlFormatter(Exception? exception)
    {
        _exception = exception;
    }

    public ErrorOpenXmlFormatter(string message)
    {
        _message = message;
    }

    public void WriteCell<TBufferWriter>(scoped ref Utf8StringWriter<TBufferWriter> writer, OpenXmlStyles styles, int rowIndex, scoped ref T value) where TBufferWriter : IBufferWriter<byte>
    {
        Throw();
    }

    public void WriteColumns<TBufferWriter>(scoped ref Utf8StringWriter<TBufferWriter> writer) where TBufferWriter : IBufferWriter<byte>
    {
        Throw();
    }

    public void WriteHeaders<TBufferWriter>(scoped ref Utf8StringWriter<TBufferWriter> writer, OpenXmlStyles styles) where TBufferWriter : IBufferWriter<byte>
    {
        Throw();
    }

    public int ColumnCount => 1;
    public string? SheetName => default;
    public bool FreezeHeader => false;
    public bool AutoFilter => false;

    [DoesNotReturn]
    private void Throw()
    {
        if (_exception != null)
        {
            OpenXmlFormatterException.ThrowRegisterInProviderFailed(typeof(T), _exception);
        }

        if (_message != null)
        {
            OpenXmlFormatterException.ThrowMessage(_message);
        }

        OpenXmlFormatterException.ThrowNotRegisteredInProvider(typeof(T));
    }
}