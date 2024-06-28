using FastExcelSlim;

var demo = new DemoEntity
{
    Name = "Demo",
    BirthDay = DateTime.Now,
    Gender = Gender.Male,
    Id = 1,
    Number = "No.1",
    LastOnline = DateTime.MinValue
};

new[] { demo }.SaveToExcel("demo.xlsx");


var simples = new List<Simple>
{
    new()
    {
        Birthday = new DateTime(2000, 2, 5),
        Name = "Jack",
        Gender = Gender.Male
    },
    new()
    {
        Birthday = new DateTime(2001, 12, 3),
        Name = "Anna",
        Gender = Gender.Female
    },
    new()
    {
        Birthday = new DateTime(1995, 3, 8),
        Name = "Jessie",
        Gender = Gender.Female
    },
    new()
    {
        Birthday = new DateTime(1986, 6, 5),
        Name = "Andy",
        Gender = Gender.Male
    },
};

simples.SaveToExcel("sample.xlsx", new SampleStyles());

OpenXmlFormatterProvider.Register(new ExternalTypeEntityFormatter());
var external = new List<ExternalTypeEntity>
{
    new (){ Gender = Gender.Male, Id = 1, Name = "Testing", Number = "No.1" },
    new (){ Gender = Gender.Female, Id = 2, Name = "Alpha", Number = "No.2" }
};
external.SaveToExcel("external.xlsx");
