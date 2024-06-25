using System.Buffers;
using System.Diagnostics;
using System.Reflection;
using Dapper;
//using System.IO.Pipelines;
//using System.Text;
using FastExcelSlim;
using FastExcelSlim.Extensions;
using FastExcelSlim.OpenXml;
using MySql.Data.MySqlClient;
using Utf8StringInterpolation;

var ms = new Student().GetType().GetMethods();
var m = new Student().GetType().GetMethod("RegisterFormatter",
    BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Static);

//var buffer = new byte[5];
//var ms = new MemoryStream(buffer);
//var pipe = PipeWriter.Create(ms);
//var utf8 = Utf8String.CreateWriter(pipe);
//utf8.ConvertXYToCellReference(677, 256);
//utf8.Flush();
//pipe.Complete();
//utf8.Dispose();
//var reference = Encoding.UTF8.GetString(buffer);
//Console.WriteLine(reference);

//var styles = new StyleTable();
//ms = new MemoryStream();
//pipe = PipeWriter.Create(ms);
//utf8 = Utf8String.CreateWriter(pipe);
//styles.Write(ref utf8);
//utf8.Flush();
//pipe.Complete();
//utf8.Dispose();
//var styleXml = Encoding.UTF8.GetString(ms.ToArray());
//Console.WriteLine(styleXml);


//var file = "test.xlsx";

//using var stream = File.Open(file, FileMode.Create, FileAccess.ReadWrite, FileShare.ReadWrite);

//var values = new Fixture().CreateMany<ExcelEntity>(1000);

//new OpenXmlWorkbookWriter<ExcelEntity>(stream, values).Save();

//[OpenXmlWritable]
//internal partial class ExcelEntity : IOpenXmlWritable<ExcelEntity>
//{
//    public int Id { get; set; }

//    public string? Name { get; set; }

//    public Guid GUIId { get; set; }

//    [OpenXmlProperty(ColumnName = "User Age", Width = 20)]
//    public int Age { get; set; }

//    [OpenXmlProperty(ColumnName = "User Birthday Year", Width = 10)]
//    private int _birthdayYear;

//    public DateTime Birthday { get; set; }

//    [OpenXmlIgnore]
//    public string? Class { get; set; }

//    public string? Address { get; set; }

//    public double Deposit { get; set; }

//    public long Friends { get; set; }

//    public bool IsOnline { get; set; }

//    public DateOnly LastOnline { get; set; }

//    public static int ColumnCount => 11;

//    public static string SheetName => "Demo";

//    static ExcelEntity()
//    {
//        OpenXmlFormatterProvider.Register<ExcelEntity>();
//    }

//    public static void RegisterFormatter()
//    {
//        if (!OpenXmlFormatterProvider.IsRegistered<ExcelEntity>())
//        {
//            OpenXmlFormatterProvider.Register(new OpenXmlFormatter<ExcelEntity>());
//        }
//    }

//    public static void WriteCell<TBufferWriter>(scoped ref Utf8StringWriter<TBufferWriter> writer, OpenXmlStyles<ExcelEntity> styles, int rowIndex,
//        scoped ref ExcelEntity value) where TBufferWriter : IBufferWriter<byte>
//    {
//        writer.WriteCell(styles, value.Id, rowIndex, 1, nameof(Id), ref value);
//        writer.WriteCell(styles, value.Name, rowIndex, 2, nameof(Name), ref value);
//        writer.WriteCell(styles, value.GUIId, rowIndex, 3, nameof(GUIId), ref value);
//        writer.WriteCell(styles, value.Age, rowIndex, 4, nameof(Age), ref value);
//        writer.WriteCell(styles, value.Birthday, rowIndex, 5, nameof(Birthday), ref value);
//        writer.WriteCell(styles, value.Class, rowIndex, 6, nameof(Class), ref value);
//        writer.WriteCell(styles, value.Address, rowIndex, 7, nameof(Address), ref value);
//        writer.WriteCell(styles, value.Deposit, rowIndex, 8, nameof(Deposit), ref value);
//        writer.WriteCell(styles, value.Friends, rowIndex, 9, nameof(Friends), ref value);
//        writer.WriteCell(styles, value.IsOnline, rowIndex, 10, nameof(IsOnline), ref value);
//        writer.WriteCell(styles, value.LastOnline, rowIndex, 11, nameof(LastOnline), ref value);
//    }

//    public static void WriteColumns<TBufferWriter>(scoped ref Utf8StringWriter<TBufferWriter> writer) where TBufferWriter : IBufferWriter<byte>
//    {
//        for (int i = 1; i <= ColumnCount; i++)
//        {
//            writer.WriteColumn(i, 15);
//        }
//    }

//    public static void WriteHeaders<TBufferWriter>(scoped ref Utf8StringWriter<TBufferWriter> writer, OpenXmlStyles<ExcelEntity> styles) where TBufferWriter : IBufferWriter<byte>
//    {
//        writer.WriteHeader(styles, nameof(Id), nameof(Id), 1);
//        writer.WriteHeader(styles, nameof(Name), nameof(Name), 2);
//        writer.WriteHeader(styles, nameof(GUIId), nameof(GUIId), 3);
//        writer.WriteHeader(styles, nameof(Age), nameof(Age), 4);
//        writer.WriteHeader(styles, nameof(Birthday), nameof(Birthday), 5);
//        writer.WriteHeader(styles, nameof(Class), nameof(Class), 6);
//        writer.WriteHeader(styles, nameof(Address), nameof(Address), 7);
//        writer.WriteHeader(styles, nameof(Deposit), nameof(Deposit), 8);
//        writer.WriteHeader(styles, nameof(Friends), nameof(Friends), 9);
//        writer.WriteHeader(styles, nameof(IsOnline), nameof(IsOnline), 10);
//        writer.WriteHeader(styles, nameof(LastOnline), nameof(LastOnline), 11);
//    }
//}

//var students = new Fixture().Build<Student>().Without(s => s.Id).CreateMany(300000);

Console.WriteLine("press enter to start");
Console.ReadLine();
Console.WriteLine("start");
const string connectionString = "Database=dev;Data Source=dev0010001;User Id=dev;Password=DEV@user12345;pooling=false;CharSet=utf8;port=3306;";
using var dbConnection = new MySqlConnection(connectionString);
dbConnection.Open();
//dbConnection.Execute("insert into Students(Name,Number,Gender) values (@Name,@Number,@Gender);", students);

var students = dbConnection.Query<Student>(new CommandDefinition("select * from Students", flags: CommandFlags.NoCache));

var file = "students.xlsx";

var stream = File.Open(file, FileMode.Create, FileAccess.ReadWrite, FileShare.ReadWrite);

var sw = new Stopwatch();
sw.Start();
new OpenXmlWorkbookWriter<Student>(stream, students).Save();
stream.Dispose();
GC.Collect();
sw.Stop();
Console.WriteLine($"done cost: {sw.Elapsed.TotalSeconds}s");
Console.ReadLine();

public class Student : IOpenXmlWritable<Student>
{
    public int Id { get; set; }

    public string Name { get; set; }

    public string Number { get; set; }

    public string Gender { get; set; }

    public static int ColumnCount => 4;

    public static string SheetName => "Student";

    static Student()
    {
        OpenXmlFormatterProvider.Register<Student>();
    }

    public static void RegisterFormatter()
    {
        if (!OpenXmlFormatterProvider.IsRegistered<Student>())
        {
            OpenXmlFormatterProvider.Register(new OpenXmlFormatter<Student>());
        }
    }

    public static void WriteCell<TBufferWriter>(scoped ref Utf8StringWriter<TBufferWriter> writer, OpenXmlStyles<Student> styles, int rowIndex,
        scoped ref Student value) where TBufferWriter : IBufferWriter<byte>
    {
        writer.WriteCell(styles, value.Id, rowIndex, 1, nameof(Id), ref value);
        writer.WriteCell(styles, value.Name, rowIndex, 2, nameof(Name), ref value);
        writer.WriteCell(styles, value.Number, rowIndex, 3, nameof(Number), ref value);
        writer.WriteCell(styles, value.Gender, rowIndex, 4, nameof(Gender), ref value);
    }

    public static void WriteColumns<TBufferWriter>(scoped ref Utf8StringWriter<TBufferWriter> writer) where TBufferWriter : IBufferWriter<byte>
    {
        for (int i = 1; i <= ColumnCount; i++)
        {
            writer.WriteColumn(i, 15);
        }
    }

    public static void WriteHeaders<TBufferWriter>(scoped ref Utf8StringWriter<TBufferWriter> writer, OpenXmlStyles<Student> styles) where TBufferWriter : IBufferWriter<byte>
    {
        writer.WriteHeader(styles, nameof(Id), nameof(Id), 1);
        writer.WriteHeader(styles, nameof(Name), nameof(Name), 2);
        writer.WriteHeader(styles, nameof(Number), nameof(Number), 3);
        writer.WriteHeader(styles, nameof(Gender), nameof(Gender), 4);
    }
}