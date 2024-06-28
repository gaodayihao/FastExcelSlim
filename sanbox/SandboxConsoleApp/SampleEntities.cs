using System.Buffers;
using FastExcelSlim.Extensions;
using Utf8StringInterpolation;

namespace FastExcelSlim;

public enum Gender
{
    Male = 1,
    Female,
    Other
}

[OpenXmlWritable(SheetName = "Demo")]
public partial class DemoEntity
{
    [OpenXmlOrder(0)]
    public int Id { get; set; }

    [OpenXmlOrder(2)]
    public string? Name { get; set; }

    [OpenXmlOrder(1)]
    [OpenXmlProperty(ColumnName = "Student Number", Width = 50)]
    public string? Number { get; set; }

    [OpenXmlOrder(3)]
    [OpenXmlProperty(ColumnName = "Gender", Width = 20)]
    [OpenXmlEnumFormat("G")]
    public Gender Gender { get; set; }

    [OpenXmlOrder(4)]
    [OpenXmlProperty(ColumnName = "Gender Value", Width = 20.66)]
    [OpenXmlEnumFormat("D")]
    private Gender NumberFormatGender => Gender;

    [OpenXmlOrder(5)]
    [OpenXmlProperty(ColumnName = "Age")]
    private int? _age = 5;

    [OpenXmlOrder(6)]
    [OpenXmlProperty(Width = 35)]
    public DateTime BirthDay { get; set; }

    [OpenXmlIgnore]
    public DateTime LastOnline { get; set; }
}

[OpenXmlWritable(FreezeHeader = true, AutoFilter = true)]
public partial class Student
{
    public int Id { get; set; }

    public string? Name { get; set; }

    public string? Number { get; set; }

    public string? Gender { get; set; }
}

[OpenXmlWritable]
internal partial class ExcelEntity
{
    public int Id { get; set; }

    public string? Name { get; set; }

    public Guid GUIId { get; set; }

    public int Age { get; set; }

    public Gender Gender { get; set; }

    public DateTime Birthday { get; set; }

    public string? Class { get; set; }

    public string? Address { get; set; }

    public double Deposit { get; set; }

    public long Friends { get; set; }

    public bool IsOnline { get; set; }

    public DateOnly LastOnline { get; set; }
}

public sealed class ExternalTypeEntity
{
    public int Id { get; set; }

    public string? Name { get; set; }

    public string? Number { get; set; }

    public Gender Gender { get; set; }
}

public class ExternalTypeEntityFormatter : IOpenXmlFormatter<ExternalTypeEntity>
{
    public void WriteCell<TBufferWriter>(scoped ref Utf8StringWriter<TBufferWriter> writer, OpenXmlStyles styles, int rowIndex,
        scoped ref ExternalTypeEntity value) where TBufferWriter : IBufferWriter<byte>
    {
        writer.WriteCell(styles, value.Id, rowIndex, 1, nameof(ExternalTypeEntity.Id), ref value);
        writer.WriteCell(styles, value.Name, rowIndex, 2, nameof(ExternalTypeEntity.Name), ref value);
        writer.WriteCell(styles, value.Number, rowIndex, 3, nameof(ExternalTypeEntity.Number), ref value);
        writer.WriteCell(styles, value.Gender, rowIndex, 4, nameof(ExternalTypeEntity.Gender), ref value);
    }

    public void WriteColumns<TBufferWriter>(scoped ref Utf8StringWriter<TBufferWriter> writer) where TBufferWriter : IBufferWriter<byte>
    {
        double[] columnWidth = [15, 20, 20, 18];
        for (int i = 0; i < columnWidth.Length; i++)
        {
            writer.WriteColumn(i + 1, columnWidth[i]);
        }
    }

    public void WriteHeaders<TBufferWriter>(scoped ref Utf8StringWriter<TBufferWriter> writer, OpenXmlStyles styles) where TBufferWriter : IBufferWriter<byte>
    {
        writer.WriteHeader(styles, "Id", nameof(ExternalTypeEntity.Id), 1);
        writer.WriteHeader(styles, "Name", nameof(ExternalTypeEntity.Name), 2);
        writer.WriteHeader(styles, "NUMBER", nameof(ExternalTypeEntity.Number), 3);
        writer.WriteHeader(styles, "GENDER", nameof(ExternalTypeEntity.Gender), 4);
    }

    public int ColumnCount => 4;

    public string? SheetName => "External";

    public bool FreezeHeader => true;

    public bool AutoFilter => false;
}