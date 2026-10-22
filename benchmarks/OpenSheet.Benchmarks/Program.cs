using BenchmarkDotNet.Running;

namespace OpenSheet.Benchmarks
{
    public static class Program
    {
        // dotnet run -c Release -f net10.0 -- --inprocess --filter "*Export*" --job short
        public static void Main(string[] args) => BenchmarkSwitcher.FromAssembly(typeof(Program).Assembly).Run(args);
    }
}
