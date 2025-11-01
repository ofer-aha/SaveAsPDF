using Microsoft.VisualStudio.TestTools.UnitTesting;
using SaveAsPDF;
using SaveAsPDF.Models;
using System;
using System.Collections.Generic;
using System.Reflection;
using System.Windows.Forms;

namespace SaveAsPDF.Tests
{
 [TestClass]
 public class FormSettingsTests
 {
 private class DummyRequester : ISettingsRequester
 {
 public void SettingsComplete(SettingsModel model) { }
 }

 private MethodInfo GetPrivateMethod(object instance, string name)
 {
 var mi = instance.GetType().GetMethod(name, BindingFlags.Instance | BindingFlags.NonPublic);
 if (mi == null) throw new InvalidOperationException($"Method '{name}' not found on type {instance.GetType().FullName}");
 return mi;
 }

 [TestMethod]
 public void LoadTreeViewFromList_ParsesIndentedLines()
 {
 // Arrange
 var fs = (FormSettings)Activator.CreateInstance(typeof(FormSettings), new object[] { new DummyRequester() });
 var mi = GetPrivateMethod(fs, "LoadTreeViewFromList");
 var tv = new TreeView();

 string[] lines = new[]
 {
 "Root1",
 " Child1",
 " Grand1",
 "Root2",
 " Child2"
 };

 // Act
 mi.Invoke(fs, new object[] { tv, lines });

 // Assert
 Assert.AreEqual(2, tv.Nodes.Count, "There should be two root nodes");
 Assert.AreEqual("Root1", tv.Nodes[0].Text);
 Assert.AreEqual("Root2", tv.Nodes[1].Text);

 Assert.AreEqual(1, tv.Nodes[0].Nodes.Count, "Root1 should have one child");
 Assert.AreEqual("Child1", tv.Nodes[0].Nodes[0].Text);
 Assert.AreEqual(1, tv.Nodes[0].Nodes[0].Nodes.Count, "Child1 should have one child");
 Assert.AreEqual("Grand1", tv.Nodes[0].Nodes[0].Nodes[0].Text);

 Assert.AreEqual(1, tv.Nodes[1].Nodes.Count, "Root2 should have one child");
 Assert.AreEqual("Child2", tv.Nodes[1].Nodes[0].Text);
 }

 [TestMethod]
 public void PopulateComboBoxFromTree_AddsFullPaths()
 {
 // Arrange
 var fs = (FormSettings)Activator.CreateInstance(typeof(FormSettings), new object[] { new DummyRequester() });
 var loadMi = GetPrivateMethod(fs, "LoadTreeViewFromList");
 var popMi = GetPrivateMethod(fs, "PopulateComboBoxFromTree");

 var tv = new TreeView();
 string[] lines = new[]
 {
 "RootA",
 " ChildA1",
 " GrandA1",
 "RootB",
 " ChildB1"
 };

 // Build tree
 loadMi.Invoke(fs, new object[] { tv, lines });

 var combo = new ComboBox();

 // Act
 popMi.Invoke(fs, new object[] { tv, combo });

 // Assert expected full paths (TreeView.PathSeparator is '\\')
 var expected = new List<string>
 {
 "RootA",
 "RootA\\ChildA1",
 "RootA\\ChildA1\\GrandA1",
 "RootB",
 "RootB\\ChildB1"
 };

 CollectionAssert.AreEqual(expected, new List<string>(GetItemsAsStrings(combo)));
 }

 private static string[] GetItemsAsStrings(ComboBox combo)
 {
 var list = new string[combo.Items.Count];
 for (int i =0; i < combo.Items.Count; i++) list[i] = combo.Items[i].ToString();
 return list;
 }
 }
}
