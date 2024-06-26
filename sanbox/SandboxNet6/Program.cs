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