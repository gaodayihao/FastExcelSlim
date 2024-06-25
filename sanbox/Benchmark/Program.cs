using BenchmarkDotNet.Running;
using FastExcelSlim.Benchmarks;

BenchmarkRunner.Run<Benchmark>();

Console.ReadLine();