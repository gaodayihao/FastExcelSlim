using System.Diagnostics;
using Dapper;
using FastExcelSlim;
using MySql.Data.MySqlClient;

Console.WriteLine("press enter to start");
Console.ReadLine();
Console.WriteLine("start");
const string connectionString = "Database=dev;Data Source=dev0010001;User Id=dev;Password=DEV@user12345;pooling=false;CharSet=utf8;port=3306;";
using var dbConnection = new MySqlConnection(connectionString);
dbConnection.Open();
//dbConnection.Execute("insert into Students(Name,Number,Gender) values (@Name,@Number,@Gender);", students);

var students = dbConnection.Query<Student>(new CommandDefinition("select * from Students", flags: CommandFlags.NoCache));

var sw = new Stopwatch();
sw.Start();
students.SaveAs("students.xlsx");
GC.Collect();
sw.Stop();
Console.WriteLine($"done cost: {sw.Elapsed.TotalSeconds}s");
Console.ReadLine();

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
    private int? _age;

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