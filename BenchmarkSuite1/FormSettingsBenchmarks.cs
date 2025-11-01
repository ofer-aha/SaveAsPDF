using BenchmarkDotNet.Attributes;
using SaveAsPDF;
using SaveAsPDF.Models;
using System;
using System.Collections.Generic;
using System.Reflection;
using System.Windows.Forms;
using Microsoft.VSDiagnostics;

namespace BenchmarkSuite1
{
    [CPUUsageDiagnoser]
    public class FormSettingsBenchmarks
    {
        private string[] _lines;
        private FormSettings _formSettingsInstance;
        private MethodInfo _loadMethod;
        private TreeView _treeView;
        private class DummyRequester : ISettingsRequester
        {
            public void SettingsComplete(SettingsModel model)
            {
            }
        }

        [GlobalSetup]
        public void Setup()
        {
            // Prepare a representative large input
            var list = new List<string>(20000);
            for (int i = 0; i < 20000; i++)
            {
                int indent = i % 6;
                list.Add(new string (' ', indent * 2) + "Folder_" + i);
            }

            _lines = list.ToArray();
            // Construct a real FormSettings instance with a dummy requester
            _formSettingsInstance = (FormSettings)Activator.CreateInstance(typeof(FormSettings), new object[] { new DummyRequester() });
            // Obtain the private method via reflection
            _loadMethod = typeof(FormSettings).GetMethod("LoadTreeViewFromList", BindingFlags.NonPublic | BindingFlags.Instance);
            if (_loadMethod == null)
                throw new InvalidOperationException("Could not find LoadTreeViewFromList via reflection.");
            // Create a TreeView to pass into the method
            _treeView = new TreeView();
        }

        [Benchmark(Description = "LoadTreeViewFromList (large dataset)")]
        public object LoadTreeViewFromList_Large()
        {
            // Invoke the private method
            _loadMethod.Invoke(_formSettingsInstance, new object[] { _treeView, _lines });
            int nodeCount = _treeView.Nodes.Count;
            _treeView.Nodes.Clear();
            return (object)nodeCount;
        }
    }
}