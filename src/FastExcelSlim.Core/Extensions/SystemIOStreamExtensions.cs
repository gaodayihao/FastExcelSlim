#if NETSTANDARD2_0
using System.Buffers;
namespace System.IO;


internal static class SystemIOStreamExtensions
{
    public static async Task WriteAsync(this Stream stream, ReadOnlyMemory<byte> data, CancellationToken cancellationToken = default)
    {
        var buffer = ArrayPool<byte>.Shared.Rent(data.Length);
        try
        {
            data.CopyTo(buffer);
            await stream.WriteAsync(buffer, 0, data.Length, cancellationToken);
        }
        finally
        {
            ArrayPool<byte>.Shared.Return(buffer);
        }
    }

    public static void Write(this Stream stream, ReadOnlyMemory<byte> data)
    {
        var buffer = ArrayPool<byte>.Shared.Rent(data.Length);
        try
        {
            data.CopyTo(buffer);
            stream.Write(buffer, 0, data.Length);
        }
        finally
        {
            ArrayPool<byte>.Shared.Return(buffer);
        }
    }
}

#endif