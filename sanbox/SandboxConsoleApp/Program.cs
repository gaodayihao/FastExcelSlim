using AutoFixture;
using FastExcelSlim;

var fixture = new Fixture();
var demos = fixture.Build<DemoEntity>().CreateMany(1000);
var students = fixture.Build<Student>().CreateMany(5000);

var stream = File.Open("sandbox.xlsx", FileMode.Create, FileAccess.ReadWrite, FileShare.ReadWrite);
stream.SaveToExcel(demos, students);
stream.Dispose();

demos = fixture.Build<DemoEntity>().CreateMany(1000);
demos.SaveToExcel("singleSheet.xlsx");

//Console.WriteLine("press enter to start");
//Console.ReadLine();
//Console.WriteLine("start");
//const string connectionString = "Database=dev;Data Source=dev0010001;User Id=dev;Password=Password;pooling=false;CharSet=utf8;port=3306;";
//using var dbConnection = new MySqlConnection(connectionString);
//dbConnection.Open();
////dbConnection.Execute("insert into Students(Name,Number,Gender) values (@Name,@Number,@Gender);", students);

//var students = dbConnection.Query<Student>(new CommandDefinition("select * from Students", flags: CommandFlags.NoCache));

//var sw = new Stopwatch();
//sw.Start();
//students.SaveToExcel("students.xlsx");
//GC.Collect();
//sw.Stop();
//Console.WriteLine($"done cost: {sw.Elapsed.TotalSeconds}s");
//Console.ReadLine();

[OpenXmlWritable("Demo")]
public partial class DemoEntity
{
    [OpenXmlOrder(0)]
    public int Id { get; set; }

    [OpenXmlOrder(2)]
    public string? Name { get; set; }

    [OpenXmlOrder(1)]
    [OpenXmlProperty("Student Number", 50)]
    public string? Number { get; set; }

    [OpenXmlOrder(3)]
    [OpenXmlProperty("Gender", 20)]
    public string? Gender { get; set; }

    [OpenXmlOrder(4)]
    [OpenXmlProperty("Age")]
    private int? _age = 5;

    [OpenXmlOrder(5)]
    [OpenXmlProperty(35)]
    public DateTime BirthDay { get; set; }

    [OpenXmlIgnore]
    public DateTime LastOnline { get; set; }
}

[OpenXmlWritable]
public partial class Student
{
    public int Id { get; set; }

    public string? Name { get; set; }

    public string? Number { get; set; }

    public string? Gender { get; set; }
}