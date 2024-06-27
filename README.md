# FastExcelSlim

![](https://img.shields.io/github/repo-size/gaodayihao/FastExcelSlim?color=green)
![](https://img.shields.io/github/languages/top/gaodayihao/FastExcelSlim)
![](https://img.shields.io/github/license/gaodayihao/FastExcelSlim)
![](https://img.shields.io/github/v/tag/gaodayihao/FastExcelSlim)
![](https://img.shields.io/github/issues/gaodayihao/FastExcelSlim)
![](https://img.shields.io/github/stars/gaodayihao/FastExcelSlim?style=social)

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

Supported field/property types
---


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

Benchmark
---
BenchmarkDotNet v0.13.12, Windows 11 (10.0.22621.3737/22H2/2022Update/SunValley2)

AMD Ryzen 7 5800H with Radeon Graphics, 1 CPU, 16 logical and 8 physical cores

.NET SDK 8.0.300
  - [Host]     : .NET 8.0.5 (8.0.524.21615), X64 RyuJIT AVX2
  - DefaultJob : .NET 8.0.5 (8.0.524.21615), X64 RyuJIT AVX2


| Method        | N     | Mean         | Error      | StdDev     | Ratio | RatioSD | Gen0        | Gen1       | Gen2      | Allocated | Alloc Ratio |
|-------------- |------ |-------------:|-----------:|-----------:|------:|--------:|------------:|-----------:|----------:|----------:|------------:|
| FastExcelSlim | 1000  |     7.078 ms |  0.0371 ms |  0.0310 ms |  0.55 |    0.01 |    484.3750 |   484.3750 |  484.3750 |   2.01 MB |        0.09 |
| MiniExcel     | 1000  |    12.949 ms |  0.2304 ms |  0.2042 ms |  1.00 |    0.00 |   2062.5000 |  1625.0000 | 1531.2500 |   22.4 MB |        1.00 |
| NPOI          | 1000  |    40.196 ms |  0.7904 ms |  0.9996 ms |  3.11 |    0.09 |   2600.0000 |  1400.0000 |  800.0000 |  20.31 MB |        0.91 |
|             - |       |              |            |            |       |         |             |            |           |           |             |
| FastExcelSlim | 10000 |    68.716 ms |  0.4566 ms |  0.3813 ms |  0.79 |    0.01 |    250.0000 |   250.0000 |  250.0000 |     16 MB |        0.24 |
| MiniExcel     | 10000 |    86.852 ms |  1.4841 ms |  1.2393 ms |  1.00 |    0.00 |   7250.0000 |  1250.0000 | 1250.0000 |     66 MB |        1.00 |
| NPOI          | 10000 |   433.318 ms |  8.5108 ms |  7.5446 ms |  5.00 |    0.12 |  21000.0000 |  8000.0000 | 1000.0000 | 204.23 MB |        3.09 |
|             - |       |              |            |            |       |         |             |            |           |           |             |
| FastExcelSlim | 50000 |   352.614 ms |  3.9816 ms |  3.7244 ms |  0.85 |    0.01 |           - |          - |         - |  64.21 MB |        0.24 |
| MiniExcel     | 50000 |   413.427 ms |  5.9601 ms |  5.2834 ms |  1.00 |    0.00 |  31000.0000 |          - |         - | 265.95 MB |        1.00 |
| NPOI          | 50000 | 2,437.711 ms | 31.0766 ms | 29.0691 ms |  5.89 |    0.11 | 108000.0000 | 34000.0000 | 4000.0000 | 996.55 MB |        3.75 |