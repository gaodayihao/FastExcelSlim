using System.Diagnostics;
using AutoFixture;
using FastExcelSlim;

Console.WriteLine("press enter to start");
Console.ReadLine();

Console.WriteLine("start fixture");
var fixture = new Fixture();
var excelEntities = fixture.Build<ExcelEntity>().CreateMany(50000);

Console.WriteLine("start export");
var stream = File.Open("excelEntities.xlsx", FileMode.Create, FileAccess.ReadWrite, FileShare.ReadWrite);
var sw = new Stopwatch();
sw.Start();
stream.WriteAsExcel(excelEntities);
stream.Dispose();
sw.Stop();
Console.WriteLine($"excel entities export done cost: {sw.Elapsed.TotalSeconds}s");

OpenXmlFormatterProvider.Register(new ExternalTypeEntityFormatter());
var external = new List<ExternalTypeEntity>
{
    new (){ Gender = Gender.Male, Id = 1, Name = "Testing", Number = "No.1" },
    new (){ Gender = Gender.Female, Id = 2, Name = "Alpha", Number = "No.2" }
};
external.SaveToExcel("external.xlsx");

//var fixture = new Fixture();
var demos = fixture.Build<DemoEntity>().CreateMany(1000);
var students = fixture.Build<Student>().CreateMany(5000);

sw.Restart();
stream = File.Open("sandbox.xlsx", FileMode.Create, FileAccess.ReadWrite, FileShare.ReadWrite);
stream.WriteAsExcel(demos, students);
stream.Dispose();
sw.Stop();
Console.WriteLine($"sandbox export done cost: {sw.Elapsed.TotalSeconds}s");
GC.Collect();
Console.ReadLine();

//const string connectionString = "Database=dev;Data Source=dev0010001;User Id=dev;Password=Password;pooling=false;CharSet=utf8;port=3306;";
//using var dbConnection = new MySqlConnection(connectionString);
//dbConnection.Open();

//var students = dbConnection.Query<Student>(new CommandDefinition("select * from Students", flags: CommandFlags.NoCache));

//var sw = new Stopwatch();
//sw.Start();
//students.SaveToExcel("students.xlsx");
//GC.Collect();
//sw.Stop();
//Console.WriteLine($"done cost: {sw.Elapsed.TotalSeconds}s");
//Console.ReadLine();