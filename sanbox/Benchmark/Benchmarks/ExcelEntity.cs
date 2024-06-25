using System.Buffers;
using FastExcelSlim.Extensions;
using Utf8StringInterpolation;

namespace FastExcelSlim.Benchmarks;

[OpenXmlWritable]
internal partial class ExcelEntity : IOpenXmlWritable<ExcelEntity>
{
    public int Id { get; set; }

    public string? Name { get; set; }

    public Guid GUIId { get; set; }

    public int Age { get; set; }

    public DateTime Birthday { get; set; }

    public string? Class { get; set; }

    public string? Address { get; set; }

    public double Deposit { get; set; }

    public long Friends { get; set; }

    public bool IsOnline { get; set; }

    public DateOnly LastOnline { get; set; }

    public static int ColumnCount => 11;

    public static string SheetName => "Demo";

    static ExcelEntity()
    {
        OpenXmlFormatterProvider.Register<ExcelEntity>();
    }

    public static void RegisterFormatter()
    {
        if (!OpenXmlFormatterProvider.IsRegistered<ExcelEntity>())
        {
            OpenXmlFormatterProvider.Register(new OpenXmlFormatter<ExcelEntity>());
        }
    }

    public static void WriteCell<TBufferWriter>(scoped ref Utf8StringWriter<TBufferWriter> writer, OpenXmlStyles<ExcelEntity> styles, int rowIndex,
        scoped ref ExcelEntity value) where TBufferWriter : IBufferWriter<byte>
    {
        writer.WriteCell(styles, value.Id, rowIndex, 1, nameof(Id), ref value);
        writer.WriteCell(styles, value.Name, rowIndex, 2, nameof(Name), ref value);
        writer.WriteCell(styles, value.GUIId, rowIndex, 3, nameof(GUIId), ref value);
        writer.WriteCell(styles, value.Age, rowIndex, 4, nameof(Age), ref value);
        writer.WriteCell(styles, value.Birthday, rowIndex, 5, nameof(Birthday), ref value);
        writer.WriteCell(styles, value.Class, rowIndex, 6, nameof(Class), ref value);
        writer.WriteCell(styles, value.Address, rowIndex, 7, nameof(Address), ref value);
        writer.WriteCell(styles, value.Deposit, rowIndex, 8, nameof(Deposit), ref value);
        writer.WriteCell(styles, value.Friends, rowIndex, 9, nameof(Friends), ref value);
        writer.WriteCell(styles, value.IsOnline, rowIndex, 10, nameof(IsOnline), ref value);
        writer.WriteCell(styles, value.LastOnline, rowIndex, 11, nameof(LastOnline), ref value);
    }

    public static void WriteColumns<TBufferWriter>(scoped ref Utf8StringWriter<TBufferWriter> writer) where TBufferWriter : IBufferWriter<byte>
    {
        for (int i = 1; i <= ColumnCount; i++)
        {
            writer.WriteColumn(i, 15);
        }
    }

    public static void WriteHeaders<TBufferWriter>(scoped ref Utf8StringWriter<TBufferWriter> writer, OpenXmlStyles<ExcelEntity> styles) where TBufferWriter : IBufferWriter<byte>
    {
        writer.WriteHeader(styles, nameof(Id), nameof(Id), 1);
        writer.WriteHeader(styles, nameof(Name), nameof(Name), 2);
        writer.WriteHeader(styles, nameof(GUIId), nameof(GUIId), 3);
        writer.WriteHeader(styles, nameof(Age), nameof(Age), 4);
        writer.WriteHeader(styles, nameof(Birthday), nameof(Birthday), 5);
        writer.WriteHeader(styles, nameof(Class), nameof(Class), 6);
        writer.WriteHeader(styles, nameof(Address), nameof(Address), 7);
        writer.WriteHeader(styles, nameof(Deposit), nameof(Deposit), 8);
        writer.WriteHeader(styles, nameof(Friends), nameof(Friends), 9);
        writer.WriteHeader(styles, nameof(IsOnline), nameof(IsOnline), 10);
        writer.WriteHeader(styles, nameof(LastOnline), nameof(LastOnline), 11);
    }
}