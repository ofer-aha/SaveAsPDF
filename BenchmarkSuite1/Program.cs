using BenchmarkDotNet.Running;

namespace BenchmarkSuite1
{
    internal class Program
    {
        static void Main(string[] args)
        {
            // Preserve assembly-wide runner and also include explicit run for HtmlHelperBenchmark
            var _ = BenchmarkRunner.Run(typeof(Program).Assembly);
            BenchmarkRunner.Run<HtmlHelperBenchmark>();
        }
    }
}
