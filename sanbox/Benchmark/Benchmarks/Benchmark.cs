using AutoFixture;
using BenchmarkDotNet.Attributes;
using Microsoft.IO;
using MiniExcelLibs;
using MiniExcelLibs.OpenXml;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace FastExcelSlim.Benchmarks;

[MemoryDiagnoser]
public class Benchmark
{
    private static readonly RecyclableMemoryStreamManager Manager = new();

    [Params(1000, 10000, 50000)]
    public int N;

    private IEnumerable<ExcelEntity>? _entities;

    [GlobalSetup]
    public void GlobalSetup()
    {
        var random = new Random();
        _entities = new Fixture().Build<ExcelEntity>()
            .With(e => e.Birthday, DateTime.Now.AddMilliseconds(random.Next(1000, 10000)))
            .With(e => e.LastOnline, DateOnly.FromDateTime(DateTime.Now.AddMilliseconds(random.Next(1000, 10000))))
            .CreateMany(N);
    }

    [Benchmark]
    public void FastExcelSlim()
    {
        using var stream = Manager.GetStream();
        _entities!.SaveAs(stream);
    }

    [Benchmark(Baseline = true)]
    public void MiniExcel()
    {
        var config = new OpenXmlConfiguration { TableStyles = TableStyles.None, AutoFilter = false };
        using var stream = Manager.GetStream();
        stream.SaveAs(_entities, configuration: config);
    }

    [Benchmark]
    public void NPOI()
    {
        using var stream = Manager.GetStream();
        var workbook = new XSSFWorkbook();
        var sheet = workbook.CreateSheet("sheet1");
        GenerateHeaders(workbook, sheet);

        var rowIndex = 1;
        foreach (var entity in _entities!)
        {
            var columnIndex = 0;
            var row = sheet.CreateRow(rowIndex++);
            CreateCell(row, entity.Id, ref columnIndex);
            CreateCell(row, entity.Name!, ref columnIndex);
            CreateCell(row, entity.GUIId, ref columnIndex);
            CreateCell(row, entity.Age, ref columnIndex);
            CreateCell(row, entity.Birthday, ref columnIndex);
            CreateCell(row, entity.Class!, ref columnIndex);
            CreateCell(row, entity.Address!, ref columnIndex);
            CreateCell(row, entity.Deposit, ref columnIndex);
            CreateCell(row, entity.Friends, ref columnIndex);
            CreateCell(row, entity.IsOnline, ref columnIndex);
            CreateCell(row, entity.LastOnline, ref columnIndex);
        }

        workbook.Write(stream);
    }

    static void CreateCell(IRow r, string value, ref int cIndex)
    {
        var cel = r.CreateCell(cIndex++);
        cel.SetCellValue(value);
    }

    static void CreateCell(IRow r, DateTime value, ref int cIndex)
    {
        var cel = r.CreateCell(cIndex++);
        cel.SetCellValue(value);
    }

    static void CreateCell(IRow r, DateOnly value, ref int cIndex)
    {
        var cel = r.CreateCell(cIndex++);
        cel.SetCellValue(value);
    }

    static void CreateCell(IRow r, bool value, ref int cIndex)
    {
        var cel = r.CreateCell(cIndex++);
        cel.SetCellValue(value);
    }

    static void CreateCell(IRow r, Guid value, ref int cIndex)
    {
        var cel = r.CreateCell(cIndex++);
        cel.SetCellValue(value.ToString());
    }

    static void CreateCell(IRow r, int value, ref int cIndex)
    {
        var cel = r.CreateCell(cIndex++);
        cel.SetCellValue(value);
    }

    static void CreateCell(IRow r, long value, ref int cIndex)
    {
        var cel = r.CreateCell(cIndex++);
        cel.SetCellValue(value);
    }

    static void CreateCell(IRow r, double value, ref int cIndex)
    {
        var cel = r.CreateCell(cIndex++);
        cel.SetCellValue(value);
    }

    public static void GenerateHeaders(IWorkbook workbook, ISheet sheet)
    {
        var headerStyle = workbook.CreateCellStyle();
        var headerFont = workbook.CreateFont();
        headerFont.IsBold = true;
        headerStyle.SetFont(headerFont);

        void CreateHeaderCell(IRow r, string header, double width, ref int cIndex)
        {
            var cell = r.CreateCell(cIndex);
            cell.CellStyle = headerStyle;
            cell.SetCellValue(header);
            sheet.SetColumnWidth(cIndex, width * 250);
            cIndex++;
        }

        var row = sheet.CreateRow(0);
        var columnIndex = 0;

        CreateHeaderCell(row, nameof(ExcelEntity.Id), 15, ref columnIndex);
        CreateHeaderCell(row, nameof(ExcelEntity.Name), 15, ref columnIndex);
        CreateHeaderCell(row, nameof(ExcelEntity.GUIId), 15, ref columnIndex);
        CreateHeaderCell(row, nameof(ExcelEntity.Age), 15, ref columnIndex);
        CreateHeaderCell(row, nameof(ExcelEntity.Birthday), 15, ref columnIndex);
        CreateHeaderCell(row, nameof(ExcelEntity.Class), 15, ref columnIndex);
        CreateHeaderCell(row, nameof(ExcelEntity.Address), 15, ref columnIndex);
        CreateHeaderCell(row, nameof(ExcelEntity.Deposit), 15, ref columnIndex);
        CreateHeaderCell(row, nameof(ExcelEntity.Friends), 15, ref columnIndex);
        CreateHeaderCell(row, nameof(ExcelEntity.IsOnline), 15, ref columnIndex);
        CreateHeaderCell(row, nameof(ExcelEntity.LastOnline), 15, ref columnIndex);
    }
}