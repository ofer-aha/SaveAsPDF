using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace SaveAsPDF.Helpers
{
    /// <summary>
    /// Provides helper methods for working with TreeView controls, including creating, populating, and managing tree nodes.
    /// </summary>
    public static class TreeHelpers
    {
        /// <summary>
        /// Creates a TreeNode representing a directory and its subdirectories.
        /// </summary>
        /// <param name="directoryInfo">The directory information to create the node from.</param>
        /// <returns>A <see cref="TreeNode"/> representing the directory.</returns>
        public static TreeNode CreateDirectoryNode(DirectoryInfo directoryInfo)
        {
            var directoryNode = new TreeNode(directoryInfo.Name) { ImageIndex = 1 };

            try
            {
                foreach (var directory in directoryInfo.GetDirectories())
                {
                    if (!directory.Attributes.HasFlag(FileAttributes.Hidden) && !directory.Attributes.HasFlag(FileAttributes.System))
                    {
                        directoryNode.Nodes.Add(CreateDirectoryNode(directory));
                    }
                }
            }
            catch (UnauthorizedAccessException)
            {
                // Ignore directories that cannot be accessed due to permissions
            }
            catch (Exception ex)
            {
                XMessageBox.Show(
                    $"שגיאה ביצירת צומת תיקיה עבור {directoryInfo.FullName}: {ex.Message}",
                    "SaveAsPDF:CreateDirectoryNode",
                    XMessageBoxButtons.OK,
                    XMessageBoxIcon.Error,
                    XMessageAlignment.Right,
                    XMessageLanguage.Hebrew
                );
            }

            return directoryNode;
        }

        /// <summary>
        /// Retrieves a list of TreeNodes representing the subdirectories of a given directory.
        /// </summary>
        /// <param name="dir">The directory to list.</param>
        /// <param name="expanded">If true, expands the tree nodes.</param>
        /// <returns>A list of <see cref="TreeNode"/> objects representing the subdirectories.</returns>
        public static List<TreeNode> GetFolderNodes(string dir, bool expanded)
        {
            var nodes = new List<TreeNode>();

            try
            {
                foreach (var subDir in Directory.GetDirectories(dir))
                {
                    if (!subDir.IsHidden())
                    {
                        var directoryInfo = new DirectoryInfo(subDir);
                        var treeNode = new TreeNode(directoryInfo.Name)
                        {
                            Tag = directoryInfo,
                            ToolTipText = subDir
                        };

                        if (Directory.GetDirectories(subDir).Any())
                        {
                            treeNode.Nodes.Add("...");
                        }

                        if (expanded)
                        {
                            treeNode.Expand();
                        }

                        nodes.Add(treeNode);
                    }
                }
            }
            catch (Exception ex)
            {
                XMessageBox.Show(
                    $"שגיאה בקבלת צמתים של תיקיות: {ex.Message}",
                    "SaveAsPDF:GetFolderNodes",
                    XMessageBoxButtons.OK,
                    XMessageBoxIcon.Error,
                    XMessageAlignment.Right,
                    XMessageLanguage.Hebrew
                );
            }

            return nodes;
        }

        /// <summary>
        /// Determines whether a directory is hidden.
        /// </summary>
        /// <param name="sDir">The directory path to check.</param>
        /// <returns>True if the directory is hidden; otherwise, false.</returns>
        private static bool IsHidden(this string sDir)
        {
            var dir = new DirectoryInfo(sDir);
            return dir.Attributes.HasFlag(FileAttributes.Hidden);
        }

        /// <summary>
        /// Traverses a directory and creates a TreeNode representing its structure.
        /// </summary>
        /// <param name="path">The path to traverse.</param>
        /// <param name="maxDepth">The maximum depth of traversal. If 0, all subdirectories are included.</param>
        /// <returns>A <see cref="TreeNode"/> representing the directory structure.</returns>
        public static TreeNode TraverseDirectory(string path, int maxDepth)
        {
            var rootNode = new TreeNode(path);

            try
            {
                if (Directory.Exists(path))
                {
                    TraverseSubdirectories(path, rootNode, maxDepth, 0);
                    rootNode.ExpandAll();
                }
                else
                {
                    rootNode.Nodes.Clear();
                }
            }
            catch (Exception ex)
            {
                XMessageBox.Show(
                    $"שגיאה במהלך סריקת התיקיה: {ex.Message}",
                    "SaveAsPDF:TraverseDirectory",
                    XMessageBoxButtons.OK,
                    XMessageBoxIcon.Error,
                    XMessageAlignment.Right,
                    XMessageLanguage.Hebrew
                );
            }

            return rootNode;
        }

        /// <summary>
        /// Recursively traverses subdirectories and adds them as child nodes to a parent TreeNode.
        /// </summary>
        /// <param name="path">The path of the directory to traverse.</param>
        /// <param name="parentNode">The parent TreeNode to add child nodes to.</param>
        /// <param name="maxDepth">The maximum depth of traversal. If 0, all subdirectories are included.</param>
        /// <param name="currentDepth">The current depth of traversal.</param>
        private static void TraverseSubdirectories(string path, TreeNode parentNode, int maxDepth, int currentDepth)
        {
            if (maxDepth > 0 && currentDepth >= maxDepth)
                return;

            try
            {
                foreach (var subDir in Directory.GetDirectories(path))
                {
                    var directoryInfo = new DirectoryInfo(subDir);
                    if (!directoryInfo.Attributes.HasFlag(FileAttributes.Hidden))
                    {
                        var subNode = new TreeNode(directoryInfo.Name);
                        parentNode.Nodes.Add(subNode);
                        TraverseSubdirectories(subDir, subNode, maxDepth, currentDepth + 1);
                    }
                }
            }
            catch (Exception ex)
            {
                XMessageBox.Show(
                    $"שגיאה בסריקת תיקיות משנה: {ex.Message}",
                    "SaveAsPDF:TraverseSubdirectories",
                    XMessageBoxButtons.OK,
                    XMessageBoxIcon.Error,
                    XMessageAlignment.Right,
                    XMessageLanguage.Hebrew
                );
            }
        }

        /// <summary>
        /// Renames the selected TreeNode in a TreeView.
        /// </summary>
        /// <param name="treeView">The TreeView control.</param>
        /// <param name="selectedNode">The selected TreeNode to rename.</param>
        public static void RenameNode(this TreeView treeView, TreeNode selectedNode)
        {
            if (selectedNode != null && selectedNode.Parent != null)
            {
                treeView.SelectedNode = selectedNode;
                treeView.LabelEdit = true;

                if (!selectedNode.IsEditing)
                {
                    selectedNode.BeginEdit();
                }
            }
            else
            {
                XMessageBox.Show(
                    "לא ניתן לשנות את שם הצומת הראשי או שלא נבחרה צומת.",
                    "בחירה לא חוקית",
                    XMessageBoxButtons.OK,
                    XMessageBoxIcon.Warning,
                    XMessageAlignment.Right,
                    XMessageLanguage.Hebrew
                );
            }
        }

        /// <summary>
        /// Adds a new TreeNode to the selected node in a TreeView.
        /// </summary>
        /// <param name="treeView">The TreeView control.</param>
        /// <param name="selectedNode">The selected TreeNode to add a child node to.</param>
        /// <param name="label">The label for the new TreeNode.</param>
        public static void AddNode(this TreeView treeView, TreeNode selectedNode, string label = "New Folder")
        {
            if (treeView.SelectedNode != null)
            {
                var newNode = treeView.SelectedNode.Nodes.Add(label);
                treeView.SelectedNode = newNode;
                newNode.Expand();

                if (label != SettingsHelpers.dateTag)
                {
                    treeView.RenameNode(newNode);
                }
            }
        }

        /// <summary>
        /// Deletes the selected TreeNode from a TreeView.
        /// </summary>
        /// <param name="treeView">The TreeView control.</param>
        public static void DeleteNode(this TreeView treeView)
        {
            var selectedNode = treeView.SelectedNode;

            if (selectedNode?.Parent != null)
            {
                selectedNode.Remove();
            }
            else
            {
                XMessageBox.Show(
                    "לא ניתן למחוק את הצומת הראשי.",
                    "פעולה לא חוקית",
                    XMessageBoxButtons.OK,
                    XMessageBoxIcon.Warning,
                    XMessageAlignment.Right,
                    XMessageLanguage.Hebrew
                );
            }
        }

        /// <summary>
        /// Retrieves the full paths of all nodes in a TreeView.
        /// </summary>
        /// <param name="parentNode">The parent TreeNode to start from.</param>
        /// <returns>A list of full paths for all nodes.</returns>
        public static List<string> ListNodesPath(TreeNode parentNode)
        {
            var paths = new List<string>();

            if (parentNode.Nodes.Count == 0)
            {
                paths.Add(parentNode.FullPath);
            }

            foreach (TreeNode childNode in parentNode.Nodes)
            {
                paths.AddRange(ListNodesPath(childNode));
            }

            return paths;
        }

        /// <summary>
        /// Populates a ComboBox with the full paths of all nodes in a TreeView.
        /// </summary>
        /// <param name="comboBox">The ComboBox to populate.</param>
        /// <param="node">The root TreeNode to start from.</param>
        public static void PopulateComboBoxFromTree(ComboBox comboBox, TreeNode node)
        {
            comboBox.Items.Add(node.FullPath);

            foreach (TreeNode childNode in node.Nodes)
            {
                PopulateComboBoxFromTree(comboBox, childNode);
            }
        }
    }
}