using Microsoft.VisualStudio.TestTools.UnitTesting;
using SaveAsPDF;
using SaveAsPDF.Helpers;
using SaveAsPDF.Models;
using System;
using System.Collections.Generic;
using System.IO;
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

 [TestClass]
 public class ThisAddInBehaviorTests
 {
 [TestMethod]
 public void ShouldRefreshOnSelectionChange_WhenPaneHidden_ReturnsFalse()
 {
 var shouldRefresh = MainPaneBehavior.ShouldRefreshOnSelectionChange(
 hasMainControl: true,
 hasMainPane: true,
 isMainPaneVisible: false);

 Assert.IsFalse(shouldRefresh);
 }

 [TestMethod]
 public void ShouldRefreshOnSelectionChange_WhenPaneVisible_ReturnsTrue()
 {
 var shouldRefresh = MainPaneBehavior.ShouldRefreshOnSelectionChange(
 hasMainControl: true,
 hasMainPane: true,
 isMainPaneVisible: true);

 Assert.IsTrue(shouldRefresh);
 }

 [TestMethod]
 public void ToggleVisibility_TogglesState()
 {
 Assert.IsTrue(MainPaneBehavior.ToggleVisibility(false));
 Assert.IsFalse(MainPaneBehavior.ToggleVisibility(true));
 }

 [TestMethod]
 public void ShouldLoadMailItemAfterToggle_OnlyWhenOpened()
 {
 Assert.IsTrue(MainPaneBehavior.ShouldLoadMailItemAfterToggle(true));
 Assert.IsFalse(MainPaneBehavior.ShouldLoadMailItemAfterToggle(false));
 }
 }

 [TestClass]
 public class FileFoldersHelperTests
 {
 [TestMethod]
 public void ProjectFullPath_RootDriveWithoutSlash_NormalizesToRootedDrive()
 {
 var result = "1234".ProjectFullPath("J:");
 var root = Path.GetPathRoot(result.FullName);

 Assert.AreEqual(@"J:\", root, true);
 }

 [TestMethod]
 public void ProjectFullPath_ThreeDigits_Uses01PrefixFolder()
 {
 string root = Path.Combine(Path.GetTempPath(), "SaveAsPdfTests");
 var result = "123".ProjectFullPath(root);

 StringAssert.EndsWith(result.FullName.TrimEnd('\\'), Path.Combine("01", "123").TrimEnd('\\'));
 }

 [TestMethod]
 public void ProjectFullPath_FourDigits_UsesFirstTwoDigitsFolder()
 {
 string root = Path.Combine(Path.GetTempPath(), "SaveAsPdfTests");
 var result = "1234".ProjectFullPath(root);

 StringAssert.EndsWith(result.FullName.TrimEnd('\\'), Path.Combine("12", "1234").TrimEnd('\\'));
 }

 [TestMethod]
 public void ProjectFullPath_SubProject_DefaultsToNestedLayoutWhenUnknown()
 {
 string root = Path.Combine(Path.GetTempPath(), "SaveAsPdfTests");
 var result = "123-45".ProjectFullPath(root);

 StringAssert.EndsWith(result.FullName.TrimEnd('\\'), Path.Combine("01", "123", "123-45").TrimEnd('\\'));
 }

 [TestMethod]
 public void ProjectFullPath_SubProject_PrefersExistingFlatFolderWhenBaseExists()
 {
 string root = Path.Combine(Path.GetTempPath(), "SaveAsPdfTests", Guid.NewGuid().ToString("N"));
 try
 {
 Directory.CreateDirectory(Path.Combine(root, "10", "1000"));
 Directory.CreateDirectory(Path.Combine(root, "10", "1000-1"));

 var result = "1000-1".ProjectFullPath(root);
 var expected = Path.Combine(root, "10", "1000-1");

 Assert.AreEqual(expected.TrimEnd('\\'), result.FullName.TrimEnd('\\'), true);
 }
 finally
 {
 if (Directory.Exists(root))
 Directory.Delete(root, true);
 }
 }

 [TestMethod]
 public void ProjectFullPath_ThreePartProject_ResolvesUsingFirstTwoPartBaseWhenExists()
 {
 string root = Path.Combine(Path.GetTempPath(), "SaveAsPdfTests", Guid.NewGuid().ToString("N"));
 try
 {
 Directory.CreateDirectory(Path.Combine(root, "10", "1000-2", "1000-2-1"));

 var result = "1000-2-1".ProjectFullPath(root);
 var expected = Path.Combine(root, "10", "1000-2", "1000-2-1");

 Assert.AreEqual(expected.TrimEnd('\\'), result.FullName.TrimEnd('\\'), true);
 }
 finally
 {
 if (Directory.Exists(root))
 Directory.Delete(root, true);
 }
 }

 [DataTestMethod]
 [DataRow("123", "01")]
 [DataRow("123A", "01")]
 [DataRow("123-A", "01")]
 [DataRow("123-AB", "01")]
 [DataRow("123-ABC", "01")]
 [DataRow("123-A-BC", "01")]
 [DataRow("123-AB-CD", "01")]
 [DataRow("123-ABC-DE", "01")]
 [DataRow("123A-B", "01")]
 [DataRow("123A-BC", "01")]
 [DataRow("123A-BCD", "01")]
 [DataRow("123A-B-CD", "01")]
 [DataRow("123A-BC-DE", "01")]
 [DataRow("123A-BCD-EF", "01")]
 [DataRow("1234", "12")]
 [DataRow("1234A", "12")]
 [DataRow("1234-A", "12")]
 [DataRow("1234-AB", "12")]
 [DataRow("1234-ABC", "12")]
 [DataRow("1234-A-BC", "12")]
 [DataRow("1234-AB-CD", "12")]
 [DataRow("1234-ABC-DE", "12")]
 [DataRow("1234A-B", "12")]
 [DataRow("1234A-BC", "12")]
 [DataRow("1234A-BCD", "12")]
 [DataRow("1234A-B-CD", "12")]
 [DataRow("1234A-BC-DE", "12")]
 [DataRow("1234A-BCD-EF", "12")]
 public void ProjectFullPath_LegalFormats_MapToExpectedRootAndLevelFolder(string projectId, string expectedLevel1)
 {
 Assert.IsTrue(projectId.SafeProjectID(), "Test data must be valid project format.");

 var result = projectId.ProjectFullPath("J:");
 var root = Path.GetPathRoot(result.FullName);

 Assert.AreEqual(@"J:\", root, true, "Root drive must normalize to J:\\");

 string relative = result.FullName.Substring(root.Length).Trim('\\');
 var parts = relative.Split(new[] { '\\' }, StringSplitOptions.RemoveEmptyEntries);

 Assert.IsTrue(parts.Length >= 2, "Expected at least <L1>\\<Project> path segments.");
 Assert.AreEqual(expectedLevel1, parts[0], true, "Unexpected level-1 folder.");
 Assert.AreEqual(projectId, parts[parts.Length - 1], true, "Final path segment must be the project id.");
 }
 }
}
