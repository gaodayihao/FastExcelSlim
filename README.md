# FastExcelSlim

![](https://img.shields.io/github/repo-size/gaodayihao/FastExcelSlim?color=green)
![](https://img.shields.io/github/languages/top/gaodayihao/FastExcelSlim)
![](https://img.shields.io/github/license/gaodayihao/FastExcelSlim)
![](https://img.shields.io/github/v/tag/gaodayihao/FastExcelSlim)
![](https://img.shields.io/github/issues/gaodayihao/FastExcelSlim)
![](https://img.shields.io/github/stars/gaodayihao/FastExcelSlim?style=social)

Lightweight, fast, low-memory usage library exports entities to Excel and supports custom styles, auto filter, and freeze header functionality.

Additionally it's native AOT friendly. Source Generator based code generation, no Dynamic CodeGen (IL.Emit).


Installation
---


Quick Start
---
```csharp
using FastExcelSlim;

[OpenXmlWritable]
public partial class Car
{
    public int Id { get; set; }

    public string Color { get; set; }

    public string Number { get; set; }
}

var cars = new List<Car>
{
    new Car
    {
        Color = "Red",
        Id = 1,
        Number = "MU-01-25-99"
    },
    new Car
    {
        Color = "Blue",
        Id = 2,
        Number = "MU-02-35-35"
    }
};

cars.SaveToExcel("cars.xlsx");
```
<img width="345" alt="screenshot1" src="https://github.com/gaodayihao/FastExcelSlim/assets/1639892/d355efc0-133a-411f-99f5-e8227efacfeb">


Supported field/property types
---

* `Enum`
* `bool`
* `Guid`
* `DateTime`
* `byte`
* `decimal`
* `float`
* `double`
* `short`
* `int`
* `long`
* `sbyte`
* `ushort`
* `uint`
* `ulong`
* `string`
* `char`

Net7 or greater:

* `DateOnly`
* `TimeOnly`


Define `[OpenXmlWritable]` `class` / `struct` / `record` / `record struct`
---
`[OpenXmlWritable]` can annotate to any `class`, `struct`, `record`, `record struct`. 
```csharp
[OpenXmlWritable]
public partial class Sample
{
    // these types are written to excel by default
    public int PublicField;
    public readonly int PublicReadonlyField;
    public int PublicProperty { get; set; }
    public int PrivateSetPublicProperty { get; private set; }
    public int ReadOnlyPublicProperty { get; }
    public int InitProperty { get; init; }
    public required int RequiredInitProperty { get; init; }

    // these types will not be written to excel by default
    int privateProperty { get; set; }
    int privateField;
    readonly int privateReadOnlyField;

    // use [OpenXmlIgnore] to remove target of a public member
    [OpenXmlIgnore]
    public int PublicProperty2 => PublicProperty + PublicField;

    // use [OpenXmlProperty] to promote a private member to excel
    [OpenXmlProperty]
    int privateField2;
    [OpenXmlProperty]
    int privateProperty2 { get; set; }
}
```

### Custom sheet name

The default exported sheet name is `sheet[id]`, and you can specify the sheet name while annotating `OpenXmlWritableAttribute`.

```csharp
[OpenXmlWritable(SheetName = "Sample")]
public partial class Sample
{
    
}
```

### Auto filter

When annotating OpenXmlWritable, you can specify whether to enable `AutoFilter`, which is disabled by default.

```csharp
[OpenXmlWritable(AutoFilter = true)]
public partial class Sample
{
    
}
```

### Freezing header
Also you can specify whether to enable freezing the header, which is disabled by default.
```csharp
[OpenXmlWritable(FreezeHeader = true)]
public partial class Sample
{

}
```

### Custom column header name
Default column header is the name of each field/property, and you can annotate `OpenXmlPropertyAttribute` to specify the column header.
```csharp
[OpenXmlWritable(FreezeHeader = true, AutoFilter = true, SheetName = "Sample")]
public partial class Sample
{
    [OpenXmlProperty(ColumnName = "User Age")]
    private int _age;
}
```

### Custom column width
You can also set the width for each column, with a default value of 15.
```csharp
[OpenXmlWritable]
public partial class Sample
{
    [OpenXmlProperty(Width = 35)]
    private int _age;
}
```
### Column order
Default column order corresponds to the order of fields/properties in the entity. You can specify the column order by annotate `OpenXmlOrderAttribute`
```csharp
[OpenXmlWritable]
public partial class Sample
{
    [OpenXmlOrder(0)]
    public int Id {get;set;}

    [OpenXmlOrder(2)]
    public string Name {get;set;}

    [OpenXmlOrder(1)]
    public string Class {get;set;}
}
```
### Enum format
You can specify the formatting an enum by annotating `OpenXmlEnumFormatAttribute`
```csharp
[OpenXmlWritable]
public partial class Sample
{
    public int Id {get;set;}

    public string Name {get;set;}

    public string Class {get;set;}

    public Gender Gender {get;set;}

    [OpenXmlEnumFormat("D")] 
    public Gender GenderValue => Gender;
}

public enum Gender
{
    Male = 1,
    Female,
    Others
}
```
<img width="502" alt="PixPin_2024-06-27_16-41-59" src="https://github.com/gaodayihao/FastExcelSlim/assets/1639892/7fb2bb20-8f57-41f8-9a88-08b0009ec5d4">

<img width="622" alt="PixPin_2024-06-27_15-59-38" src="https://github.com/gaodayihao/FastExcelSlim/assets/1639892/6d81c5eb-cb73-4081-9a8a-c5a5080ce441">
<img width="617" alt="PixPin_2024-06-27_15-58-59" src="https://github.com/gaodayihao/FastExcelSlim/assets/1639892/b8ea0124-35f6-4875-9170-ce99393a8815">


Benchmark
---
BenchmarkDotNet v0.13.12, Windows 11 (10.0.22621.3737/22H2/2022Update/SunValley2)

AMD Ryzen 7 5800H with Radeon Graphics, 1 CPU, 16 logical and 8 physical cores

.NET SDK 8.0.300
  - [Host]     : .NET 8.0.5 (8.0.524.21615), X64 RyuJIT AVX2
  - DefaultJob : .NET 8.0.5 (8.0.524.21615), X64 RyuJIT AVX2


| Method        | N     | Mean         | StdDev     | Ratio |Gen0        | Gen1       | Gen2      | Allocated | Alloc Ratio |
|-------------- |------ |-------------:|-----------:|------:|-----------:|-----------:|----------:|----------:|------------:|
| FastExcelSlim | 1000  |     7.078 ms |  0.0310 ms |  0.55 |   484.3750 |   484.3750 |  484.3750 |   2.01 MB |        0.09 |
| MiniExcel     | 1000  |    12.949 ms |  0.2042 ms |  1.00 |  2062.5000 |  1625.0000 | 1531.2500 |   22.4 MB |        1.00 |
| NPOI          | 1000  |    40.196 ms |  0.9996 ms |  3.11 |  2600.0000 |  1400.0000 |  800.0000 |  20.31 MB |        0.91 |
|             - |       |              |            |       |            |            |           |           |             |
| FastExcelSlim | 10000 |    68.716 ms |  0.3813 ms |  0.79 |   250.0000 |   250.0000 |  250.0000 |     16 MB |        0.24 |
| MiniExcel     | 10000 |    86.852 ms |  1.2393 ms |  1.00 |  7250.0000 |  1250.0000 | 1250.0000 |     66 MB |        1.00 |
| NPOI          | 10000 |   433.318 ms |  7.5446 ms |  5.00 | 21000.0000 |  8000.0000 | 1000.0000 | 204.23 MB |        3.09 |
|             - |       |              |            |       |            |            |           |           |             |
| FastExcelSlim | 50000 |   352.614 ms |  3.7244 ms |  0.85 |          - |          - |         - |  64.21 MB |        0.24 |
| MiniExcel     | 50000 |   413.427 ms |  5.2834 ms |  1.00 | 31000.0000 |          - |         - | 265.95 MB |        1.00 |
| NPOI          | 50000 | 2,437.711 ms | 29.0691 ms |  5.89 |108000.0000 | 34000.0000 | 4000.0000 | 996.55 MB |        3.75 |
