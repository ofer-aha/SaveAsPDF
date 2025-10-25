using System;
using System.IO;
using BenchmarkDotNet.Attributes;
using SaveAsPDF.Helpers;
using Microsoft.VSDiagnostics;

namespace SaveAsPDF.Benchmarks
{
    [CPUUsageDiagnoser]
    public class TreeHelpersBenchmarks
    {
        private string _rootPath;
        private DirectoryInfo _rootDirInfo;
        [Params(3)]
        public int Depth;
        [Params(5)]
        public int Breadth;
        [GlobalSetup]
        public void Setup()
        {
            _rootPath = Path.Combine(Path.GetTempPath(), "Bench_Tree_" + Guid.NewGuid().ToString("N"));
            _rootDirInfo = new DirectoryInfo(_rootPath);
            _rootDirInfo.Create();
            CreateTree(_rootPath, Depth, Breadth, 0);
        }

        private void CreateTree(string basePath, int maxDepth, int breadth, int currentDepth)
        {
            if (currentDepth >= maxDepth)
                return;
            for (int i = 0; i < breadth; i++)
            {
                var dir = Path.Combine(basePath, $"L{currentDepth}_N{i}");
                Directory.CreateDirectory(dir);
                CreateTree(dir, maxDepth, breadth, currentDepth + 1);
            }
        }

        [GlobalCleanup]
        public void Cleanup()
        {
            try
            {
                if (Directory.Exists(_rootPath))
                {
                    Directory.Delete(_rootPath, true);
                }
            }
            catch
            {
            }
        }

        [Benchmark(Description = "CreateDirectoryNode (recursive full traversal)")]
        public object CreateDirectoryNode_Full()
        {
            return TreeHelpers.CreateDirectoryNode(_rootDirInfo);
        }

        [Benchmark(Description = "TraverseDirectory (maxDepth=1)")]
        public object TraverseDirectory_MaxDepth1()
        {
            return TreeHelpers.TraverseDirectory(_rootPath, 1);
        }
    }
}