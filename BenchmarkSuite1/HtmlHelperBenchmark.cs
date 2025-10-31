using BenchmarkDotNet.Attributes;
using SaveAsPDF.Helpers;
using SaveAsPDF.Models;
using System;
using System.Collections.Generic;
using System.IO;
using Microsoft.VSDiagnostics;

namespace BenchmarkSuite1
{
    [CPUUsageDiagnoser]
    // Benchmark for HtmlHelper.GenerateHtmlContent (now using GenerateHtmlToFile)
    public class HtmlHelperBenchmark
    {
        private string _path;
        private List<EmployeeModel> _smallEmployees;
        private List<AttachmentsModel> _smallAttachments;
        private List<EmployeeModel> _largeEmployees;
        private List<AttachmentsModel> _largeAttachments;
        private string _projectName;
        private string _projectId;
        private string _notes;
        private string _userName;
        [GlobalSetup]
        public void Setup()
        {
            // Use temp path to avoid writing into the repo
            _path = Path.GetTempPath();
            _projectName = "BenchmarkProject";
            _projectId = "12345";
            _notes = "These are benchmark notes.";
            _userName = "BenchUser";
            _smallEmployees = new List<EmployeeModel>(5);
            for (int i = 0; i < 5; i++)
            {
                _smallEmployees.Add(new EmployeeModel { FirstName = $"First{i}", LastName = $"Last{i}", EmailAddress = $"user{i}@example.com" });
            }

            _smallAttachments = new List<AttachmentsModel>(5);
            for (int i = 0; i < 5; i++)
            {
                _smallAttachments.Add(new AttachmentsModel { fileName = $"file{i}.txt", fileSize = $"{i * 10} KB" });
            }

            _largeEmployees = new List<EmployeeModel>(500);
            for (int i = 0; i < 500; i++)
            {
                _largeEmployees.Add(new EmployeeModel { FirstName = $"F{i}", LastName = $"L{i}", EmailAddress = $"u{i}@example.com" });
            }

            _largeAttachments = new List<AttachmentsModel>(200);
            for (int i = 0; i < 200; i++)
            {
                _largeAttachments.Add(new AttachmentsModel { fileName = $"bigfile{i}.pdf", fileSize = $"{i * 5} KB" });
            }
        }

        [Benchmark(Description = "GenerateHtmlContentToFile (small dataset)")]
        public object GenerateHtml_Small()
        {
            string tempFile = Path.Combine(_path, $"bench_small_{Guid.NewGuid()}.html");
            try
            {
                HtmlHelper.GenerateHtmlToFile(tempFile, _path, _smallEmployees, _smallAttachments, _projectName, _projectId, _notes, _userName);
                bool exists = File.Exists(tempFile);
                return (object)exists;
            }
            finally
            {
                try { if (File.Exists(tempFile)) File.Delete(tempFile); } catch { }
            }
        }

        [Benchmark(Description = "GenerateHtmlContentToFile (large dataset)")]
        public object GenerateHtml_Large()
        {
            string tempFile = Path.Combine(_path, $"bench_large_{Guid.NewGuid()}.html");
            try
            {
                HtmlHelper.GenerateHtmlToFile(tempFile, _path, _largeEmployees, _largeAttachments, _projectName, _projectId, _notes, _userName);
                bool exists = File.Exists(tempFile);
                return (object)exists;
            }
            finally
            {
                try { if (File.Exists(tempFile)) File.Delete(tempFile); } catch { }
            }
        }
    }
}