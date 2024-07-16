﻿using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using MessageBox = System.Windows.Forms.MessageBox;

namespace SaveAsPDF.Helpers
{
    public static class TreeHelpers
    {
        //XmlDocument xmlDocument;
        //TreeNode _mySelectedNode;

        /// <summary>
        /// List folders to treeView
        /// </summary>
        /// <param name="directoryInfo"></param>
        /// <returns>TreeNode folder name</returns>
        public static TreeNode CreateDirectoryNode(DirectoryInfo directoryInfo)
        {
            var directoryNode = new TreeNode(directoryInfo.Name) { ImageIndex = 1 };
            try
            {
                foreach (var directory in directoryInfo.GetDirectories())
                {
                    if (!directory.Attributes.HasFlag(FileAttributes.Hidden))
                    {
                        directoryNode.Nodes.Add(CreateDirectoryNode(directory));
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "SaveAsPDF:CreateDirectoryNode");
            }
            return directoryNode;
        }

        /// <summary>
        /// List folder to tree nodes
        /// </summary>
        /// <param name="dir">Directory to list</param>
        /// <param name="expanded">If True, expand the tree node</param>
        /// <returns>List of <see cref="TreeNode"/> tree nodes</returns>
        public static List<TreeNode> getFolderNodes(string dir, bool expanded)
        {
            string[] dirs = Directory.GetDirectories(dir).ToArray();
            List<TreeNode> nodes = new List<TreeNode>();
            foreach (string d in dirs)
            {
                if (!d.IsHidden())
                {
                    DirectoryInfo di = new DirectoryInfo(d);
                    TreeNode tn = new TreeNode(di.Name);
                    tn.Tag = di;
                    tn.ToolTipText = d;
                    int subCount = 0;
                    try { subCount = Directory.GetDirectories(d).Length; }
                    catch { /* ignore access denied */  }
                    if (subCount > 0)
                    {
                        tn.Nodes.Add("...");
                    }
                    if (expanded)
                    {
                        tn.Expand();   //  **
                    }
                    nodes.Add(tn);
                }
            }
            return nodes;
        }
        /// <summary>
        /// Test if a directory is hidden
        /// </summary>
        /// <param name="sDir"></param>
        /// <returns></returns>
        private static bool IsHidden(this string sDir)
        {
            DirectoryInfo dir = new DirectoryInfo(sDir);
            if (dir.Attributes.HasFlag(FileAttributes.Hidden))
            {
                return true;
            }
            return false;
        }
        /// <summary>
        /// Travers folders to tree nodes 
        /// 
        /// </summary>
        /// <param name="path"> Input path to start traversal </param>
        /// <param name="maxDepth"> how deep the traversal will go maxDepth=0 all sub folders</param>
        /// <returns>
        /// tree nodes 
        /// </returns>
        public static TreeNode TraverseDirectory(string path, int maxDepth)
        {
            TreeNode result = new TreeNode(path);

            try
            {
                if (Directory.Exists(path))
                {
                    TraverseSubdirectories(path, result, maxDepth, 0);
                    result.ExpandAll();
                }
                else
                {
                    result.Nodes.Clear();
                }
            }
            catch (Exception ex)
            {
                // Handle the exception (e.g., log it, show an error message, etc.)
                // You can customize this part based on your application's requirements.
                Console.WriteLine($"Error while traversing directory: {ex.Message}");
            }

            return result;
        }

        private static void TraverseSubdirectories(string path, TreeNode parentNode, int maxDepth, int currentDepth)
        {
            if (currentDepth >= maxDepth)
                return;

            foreach (string subdirectory in Directory.GetDirectories(path))
            {
                if (!File.GetAttributes(subdirectory).HasFlag(FileAttributes.Hidden))
                {
                    string subNodeName = Path.GetFileName(subdirectory); // Get only the directory name
                    TreeNode subNode = new TreeNode(subNodeName);
                    parentNode.Nodes.Add(subNode);
                    if (maxDepth != 0)
                    {
                        TraverseSubdirectories(subdirectory, subNode, maxDepth, currentDepth + 1);
                    }
                }
            }
        }


        //}
        /// <summary>
        /// Rename the node name according to user typing
        /// </summary>
        /// <param name="treeView"></param>
        /// <param name="mySelectedNode"></param>
        public static void RenameNode(this TreeView treeView, TreeNode mySelectedNode)
        {
            if (mySelectedNode != null & mySelectedNode.Parent != null)
            {
                treeView.SelectedNode = mySelectedNode;
                treeView.LabelEdit = true;
                if (!mySelectedNode.IsEditing)
                {
                    mySelectedNode.BeginEdit();
                }
            }
            else
            {
                MessageBox.Show("No tree node selected or selected node is a root node.\n" +
                   "Editing of root nodes is not allowed.", "Invalid selection");
            }
        }
        /// <summary>
        /// Rename the node name to 'newName' parameter
        /// </summary>
        /// <param name="treeView"></param>
        /// <param name="mySelectedNode"></param>
        /// <param name="newName"></param>
        public static void RenameNode(this TreeView treeView, TreeNode mySelectedNode, string newName)
        {
            if (mySelectedNode != null && mySelectedNode.Parent != null)
            {
                //treeView.SelectedNode = _mySelectedNode;
                treeView.SelectedNode.Name = newName.SafeFolderName();
            }
            else
            {
                MessageBox.Show("No tree node selected or selected node is a root node.\n" +
                   "Editing of root nodes is not allowed.", "Invalid selection");
            }
        }


        /// <summary>
        /// Add Node to tree-view
        /// </summary>
        /// <param name="treeView"></param>
        /// <param name="mySelectedNode"></param>
        /// <param name="label"></param>
        public static void AddNode(this TreeView treeView, TreeNode mySelectedNode, string label)
        {
            if (treeView.SelectedNode != null)
            {
                if (!string.IsNullOrEmpty(label))
                {
                    mySelectedNode = treeView.SelectedNode.Nodes.Add(label);

                    treeView.SelectedNode = mySelectedNode;
                    mySelectedNode.Expand();

                    if (label != SettingsHelpers.dateTag)
                    {
                        treeView.RenameNode(mySelectedNode);
                    }
                }
                else
                {
                    //if Label is Null add default node name 
                    treeView.AddNode(mySelectedNode);
                }
            }
        }
        /// <summary>
        /// Add node to tree with default node name "New Folder"
        /// </summary>
        /// <param name="treeView"></param>
        /// <param name="mySelectedNode"></param>
        public static void AddNode(this TreeView treeView, TreeNode mySelectedNode)
        {
            if (treeView.SelectedNode != null)
            {
                mySelectedNode = treeView.SelectedNode.Nodes.Add("New Folder");

                treeView.SelectedNode = mySelectedNode;
                mySelectedNode.Expand();

            }
        }
        /// <summary>
        /// Delete a node from tree-view 
        /// </summary>
        /// <param name="treeView"></param>
        /// <param name="mySelectedNode"></param>
        public static void DelNode(this TreeView treeView, TreeNode mySelectedNode)
        {
            TreeNode node = treeView.SelectedNode;

            if (node.Parent != null)
            {
                treeView.SelectedNode.Nodes.Remove(treeView.SelectedNode);
            }
            else
            {
                MessageBox.Show("Can not delete node: Root node");
            }
        }

        /// <summary>
        /// Return the root node's name as string
        /// </summary>
        /// <param name="tv"></param>
        /// <returns></returns>
        public static String RootNodeName(TreeView tv)
        {
            string output = "";
            foreach (TreeNode n in tv.Nodes)
            {
                if (n.Parent == null)
                {
                    output = n.Text;
                }
            }
            return output;
        }
        public static List<string> ListNodesPath(TreeNode oParentNode, ComboBox comboBox)
        {
            List<string> list = new List<string>();
            if (oParentNode.Nodes.Count == 0)
            {
                list.Add(oParentNode.FullPath);
                comboBox.Items.Add(oParentNode.FullPath);
            }
            // Start recursion on all sub-nodes.
            foreach (TreeNode oSubNode in oParentNode.Nodes)
            {
                list.AddRange(ListNodesPath(oSubNode, comboBox));
            }
            return list;
        }
        /// <summary>
        /// Convert a node to List of Strings full path  
        /// </summary>
        /// <param name="oParentNode">The TreeNode</param>
        /// <returns>List of Strings </returns>
        public static List<string> ListNodesPath(TreeNode oParentNode)
        {
            List<string> list = new List<string>();
            if (oParentNode.Nodes.Count == 0)
            {
                list.Add(oParentNode.FullPath);
            }
            // Start recursion on all sub-nodes.
            foreach (TreeNode oSubNode in oParentNode.Nodes)
            {
                list.AddRange(ListNodesPath(oSubNode));
            }
            return list;
        }

        /// <summary>
        /// Copy the folder's full path list to combo-box
        /// </summary>
        /// <param name="combo"></param>
        /// <param name="node"></param>
        public static void TvNodesToCombo(ComboBox combo, TreeNode node)
        {

            combo.Items.Add(node.FullPath);

            foreach (TreeNode actualNode in node.Nodes)
            {
                TvNodesToCombo(combo, actualNode);

            }

        }


        /// <summary>
        /// Write the TreeView's values into a file that uses tabs
        /// to show indentation.
        /// </summary>
        /// <param name="file_name"></param>
        /// <param name="trv"></param>
        public static void SaveTreeViewIntoFile(string file_name, TreeView trv)
        {
            List<string> lst = new List<string>();
            lst = GetPathsFromTreeView(trv.Nodes);
            try
            {
                using (StreamWriter sw = new StreamWriter(file_name))
                {
                    foreach (string line in lst)
                    {
                        sw.WriteLine(line);
                    }
                }
            }
            catch (Exception e)

            {
                MessageBox.Show(e.Message, "TreeHelper: SaveTreeViewIntoFile()");
            }
        }


        /// <summary>
        /// Convert the TreeView control to a list of paths.
        /// </summary>
        /// <param name="treeView">The TreeView control to convert.</param>
        /// <returns>A list of paths representing the nodes in the TreeView control.</returns>
        public static List<string> TreeToList(TreeView treeView)
        {
            List<string> list = new List<string>();
            foreach (TreeNode node in treeView.Nodes)
            {
                list.Add(node.FullPath);
            }
            return (list);
        }

        /// <summary>
        /// Load the default folder tree into the TreeView control.
        /// If the folder tree file does not exist, a default folder tree will be loaded.
        /// </summary>
        /// <param name="treeView">The TreeView control to load the folder tree into.</param>
        public static void LoadTreeViewFromList(this TreeView treeView)
        {
            string[] defaultTree = new string[]
            {
                @"_מספר_פרויקט_",
                @"_מספר_פרויקט_\DWG",
                @"_מספר_פרויקט_\מכתבים",
                @"_מספר_פרויקט_\OLD",
                @"_מספר_פרויקט_\PDF",
                @"_מספר_פרויקט_\אישור ציוד",
                @"_מספר_פרויקט_\אישור ציוד\לוח חשמל ראשי",
                @"_מספר_פרויקט_\אישור ציוד\דיזל גנרטור",
                @"_מספר_פרויקט_\אישור ציוד\גופי תאורה",
                @"_מספר_פרויקט_\התקבל",
                @"_מספר_פרויקט_\התקבל\אדריכלות",
                @"_מספר_פרויקט_\התקבל\קונסטרוקציה",
                @"_מספר_פרויקט_\התקבל\מיזוג אויר",
                @"_מספר_פרויקט_\התקבל\בטיחות",
                @"_מספר_פרויקט_\התקבל\כבישים",
                @"_מספר_פרויקט_\נשלח",
                @"_מספר_פרויקט_\נשלח\אדריכלות",
                @"_מספר_פרויקט_\נשלח\קונסטרוקציה",
                @"_מספר_פרויקט_\נשלח\מיזוג אויר",
                @"_מספר_פרויקט_\נשלח\בטיחות",
                @"_מספר_פרויקט_\נשלח\כבישים"
            };
            treeView.LoadTreeViewFromFile(defaultTree);
        }


        /// <summary>
        /// Load the TreeView control from a list of paths.
        /// </summary>
        /// <param name="treeView">The TreeView control to load the folder tree into.</param>
        /// <param name="list">The list of paths to populate the TreeView control.</param>
        public static void LoadTreeViewFromFile(this TreeView treeView, string[] list)
        {
            treeView.Nodes.Clear();
            foreach (string path in list)
            {
                string[] nodeNames = path.Split('\\');
                TreeNode parentNode = null;
                foreach (string nodeName in nodeNames)
                {
                    if (parentNode == null)
                    {
                        parentNode = treeView.Nodes.Cast<TreeNode>().FirstOrDefault(n => n.Text == nodeName);
                        if (parentNode == null)
                        {
                            parentNode = treeView.Nodes.Add(nodeName);
                        }
                    }
                    else
                    {
                        TreeNode childNode = parentNode.Nodes.Cast<TreeNode>().FirstOrDefault(n => n.Text == nodeName);
                        if (childNode == null)
                        {
                            childNode = parentNode.Nodes.Add(nodeName);
                        }
                        parentNode = childNode;
                    }
                }
            }
        }


        /// <summary>
        /// Load a TreeView control from a file that uses tabs
        /// to show indentation.
        /// </summary>
        /// <param name="file_name"></param>
        /// <param name="trv"></param>
        public static void LoadTreeViewFromFile(this TreeView trv, string file_name)
        {
            if (!string.IsNullOrEmpty(file_name))
            {
                try
                {
                    // Get the file's contents.
                    string file_contents = File.ReadAllText(file_name);
                    // Process the lines.
                    trv.Nodes.Clear();
                    Dictionary<int, TreeNode> parents = new Dictionary<int, TreeNode>();

                    // Break the file into lines.
                    string[] lines = file_contents.Split(new char[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
                    foreach (string text_line in lines)
                    {
                        // See how many tabs are at the start of the line.
                        int level = text_line.Length - text_line.TrimStart('\t').Length;

                        // Add the new node.
                        if (level == 0)
                            parents[level] = trv.Nodes.Add(text_line.Trim());
                        else
                            parents[level] = parents[level - 1].Nodes.Add(text_line.Trim());
                        parents[level].EnsureVisible();
                    }

                }
                catch (System.Exception e)
                {
                    var trace = new StackTrace(e);
                    var frame = trace.GetFrame(0);
                    var method = frame.GetMethod();
                    MessageBox.Show($"{method.DeclaringType.FullName}\n{method.Name}\n{e.Message}", "TreeHelper:LoadTreeViewFromFile()");
                }

            }
            else
            {
                trv.LoadTreeViewFromList();
                //MessageBox.Show("File name is empty", "TreeHelper:LoadTreeViewFromFile()");
            }
        }






        /// <summary>
        /// Populate the tree view 
        /// </summary>
        /// <param name="parentNode"></param>
        /// <param name="directory"></param>
        public static void PopulateTreeView(TreeNode parentNode, string directory)
        {
            try
            {
                foreach (string subdirectory in Directory.GetDirectories(directory))
                {
                    TreeNode childNode = new TreeNode(Path.GetFileName(subdirectory));
                    parentNode.Nodes.Add(childNode);
                    PopulateTreeView(childNode, subdirectory); // Recursive call for subdirectories
                }
            }
            catch (UnauthorizedAccessException)
            {
                // Handle unauthorized access (optional)
            }
        }
        /// <summary>
        ///  Populate the tree view 
        /// </summary>
        /// <param name="treeView"></param>
        /// <param name="paths"></param>
        /// <param name="pathSeparator"></param>
        public static void PopulateTreeView(TreeView treeView, IEnumerable<string> paths, char pathSeparator)
        {
            TreeNode lastNode = null;
            string subPathAgg;

            foreach (string path in paths)
            {
                subPathAgg = string.Empty;

                foreach (string subPath in path.Split(pathSeparator))
                {
                    subPathAgg += subPath + pathSeparator;
                    TreeNode[] nodes = treeView.Nodes.Find(subPathAgg, true);

                    if (nodes.Length == 0)
                    {
                        if (lastNode == null)
                            lastNode = treeView.Nodes.Add(subPathAgg, subPath);
                        else
                            lastNode = lastNode.Nodes.Add(subPathAgg, subPath);
                    }
                    else
                    {
                        lastNode = nodes[0];
                    }
                }
            }
        }
        /// <summary>
        /// Find the specific node by path 
        /// </summary>
        /// <param name="nodes"></param>
        /// <param name="path"></param>
        /// <returns></returns>
        public static TreeNode FindNodeByPath(TreeNodeCollection nodes, string path)
        {
            string[] nodeNames = path.Split('\\');
            TreeNode currentNode = null;
            foreach (string nodeName in nodeNames)
            {
                currentNode = nodes.Cast<TreeNode>().FirstOrDefault(n => n.Text == nodeName);
                if (currentNode == null)
                {
                    break;
                }
                nodes = currentNode.Nodes;
            }
            return currentNode;
        }
        /// <summary>
        /// Retrieves the paths of all nodes in a TreeView.
        /// </summary>
        /// <param name="nodes">The collection of TreeNodes to retrieve paths from.</param>
        /// <param name="path">The current path of the nodes.</param>
        /// <returns>A list of paths for all nodes in the TreeView.</returns>
        static List<string> GetPathsFromTreeView(TreeNodeCollection nodes, string path = "")
        {
            var paths = new List<string>();

            foreach (TreeNode node in nodes)
            {
                string currentPath = path + node.Text;

                if (node.Nodes.Count > 0)
                {
                    paths.AddRange(GetPathsFromTreeView(node.Nodes, currentPath + @"\"));
                }
                else
                {
                    paths.Add(currentPath);
                }
            }

            return paths;
        }
    }
}
