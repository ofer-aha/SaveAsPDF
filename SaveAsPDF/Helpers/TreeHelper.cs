using SaveAsPDF.Properties;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Windows.Forms;

namespace SaveAsPDF.Helpers
{
    public static class TreeHelper
    {
        //XmlDocument xmlDocument;
        //TreeNode mySelectedNode;
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
        /// <returns>list of tree nodes</returns>
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
        public static TreeNode TraverseDirectory(string path)
        {
            TreeNode result = new TreeNode(path);
            Cursor.Current = Cursors.WaitCursor; //show the user we are doing something..... "the computer is thinking"
            if (Directory.Exists(path))
            {
                foreach (string subdirectory in Directory.GetDirectories(path))
                {
                    if (!subdirectory.IsHidden())
                    {
                        TraverseDirectory(subdirectory);
                        string[] s = subdirectory.Split('\\');
                        result.Nodes.Add(s[s.Length - 1]);

                        //result.Nodes.Add(TraverseDirectory(subdirectory);
                    }
                }
                result.ExpandAll();
                return result;
            }
            else
            {
                result.Nodes.Clear();
                Cursor.Current = Cursors.Default;
                return result;
            }
        }
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
                //treeView.SelectedNode = mySelectedNode;
                treeView.SelectedNode.Name = newName.SafeFolderName();
                //if (!mySelectedNode.IsEditing)
                //{
                //    mySelectedNode.BeginEdit();
                //}
            }
            else
            {
                MessageBox.Show("No tree node selected or selected node is a root node.\n" +
                   "Editing of root nodes is not allowed.", "Invalid selection");
            }
        }


        /// <summary>
        /// Add Node to treeview
        /// </summary>
        /// <param name="treeView"></param>
        /// <param name="mySelectedNode"></param>
        /// <param name="lable"></param>
        public static void AddNode(this TreeView treeView, TreeNode mySelectedNode, string lable)
        {
            if (treeView.SelectedNode != null)
            {
                if (!string.IsNullOrEmpty(lable))
                {
                    mySelectedNode = treeView.SelectedNode.Nodes.Add(lable);

                    treeView.SelectedNode = mySelectedNode;
                    mySelectedNode.Expand();

                    if (lable != Settings.Default.dateTag)
                    {
                        treeView.RenameNode(mySelectedNode);
                    }
                }
                else
                {
                    //if Lable is Null add default node name 
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

        public static void DelNode(this TreeView treeView, TreeNode mySelectedNode)
        {
            TreeNode node = treeView.SelectedNode;

            if (node.Parent != null)
            {
                treeView.SelectedNode.Nodes.Remove(treeView.SelectedNode);
            }
            else
            {
                MessageBox.Show("Root node");
            }

        }

        /// <summary>
        /// Return the root node's name as strig
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

            // Start recursion on all subnodes.
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
        /// Write this node's subtree into the StringBuilder.
        /// </summary>
        /// <param name="level">Node Level as int.</param>
        /// <param name="node">Node object</param>
        /// <param name="sb">string line to be processed</param>
        private static void WriteNodeIntoString(int level, TreeNode node, StringBuilder sb)
        {
            // Append the correct number of tabs and the node's text.
            sb.AppendLine(new string('\t', level) + node.Text);

            // Recursively add children with one greater level of tabs.
            foreach (TreeNode child in node.Nodes)
                WriteNodeIntoString(level + 1, child, sb);
        }



        /// <summary>
        /// Write the TreeView's values into a file that uses tabs
        /// to show indentation.
        /// </summary>
        /// <param name="file_name"></param>
        /// <param name="trv"></param>
        public static void SaveTreeViewIntoFile(string file_name, TreeView trv)
        {
            // Build a string containing the TreeView's contents.
            StringBuilder sb = new StringBuilder();
            foreach (TreeNode node in trv.Nodes)
            {
                WriteNodeIntoString(0, node, sb);
            }
            // Write the result into the file.
            File.WriteAllText(file_name, sb.ToString());
        }


        /// <summary>
        /// 
        /// </summary>
        /// <param name="treeView"></param>
        /// <returns></returns>
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
        /// Load a TreeView control from a file that uses tabs
        /// to show indentation.
        /// </summary>
        /// <param name="file_name"></param>
        /// <param name="trv"></param>
        public static void LoadTreeViewFromFile(this TreeView trv, string file_name)
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


        /// <summary>
        /// Load the default folder tree from the default folder free file.
        /// If the folder tree file does not exist a default folder tree will be loaded
        /// </summary>
        /// <param name="treeView"></param>
        public static void LoadDefaultTree(this TreeView treeView)
        {
            Assembly assembly = Assembly.GetExecutingAssembly();
            string location = assembly.CodeBase;
            string fullPath = new Uri(location).LocalPath; // path including the dll 
            string directoryPath = Path.GetDirectoryName(fullPath); // directory path 

            string defaultTreeFileName = directoryPath + "\\" + Settings.Default.defaultTreeFile;


            if (File.Exists(defaultTreeFileName))
            {
                treeView.LoadTreeViewFromFile(defaultTreeFileName);
            }
            else
            {
                //Default tree file was not found -  create a new default free file
                string defaultTree = "_מספר_פרויקט_\n" +
                    "\tDWG\n" +
                    "\tמכתבים\n" +
                    "\tOLD\n" +
                    "\tPDF\n" +
                    "\tאישור ציוד\n" +
                    "\t\tלוח חשמל ראשי\n" +
                    "\t\tדיזל גנרטור\n" +
                    "\t\tגופי תאורה\n" +
                    "\tהתקבל\n" +
                    "\t\tאדריכלות\n" +
                    "\t\tקונסטרוקציה\n" +
                    "\t\tמיזוג אויר\n" +
                    "\t\tבטיחות\n" +
                    "\t\tכבישים\n" +
                    "\tנשלח\n" +
                    "\t\tאדריכלות\n" +
                    "\t\tקונסטרוקציה\n" +
                    "\t\tמיזוג אויר\n" +
                    "\t\tבטיחות\n" +
                    "\t\tכבישים\n ";

                File.WriteAllText(defaultTreeFileName, defaultTree);
                treeView.LoadTreeViewFromFile(defaultTreeFileName);

            }
        }
        /// <summary>
        /// Load folder path to tree-view
        /// </summary>
        /// <param name="tv"></param>
        /// <param name="folderPath"></param>
        public static void LoadFolderPaths(TreeView tv, string folderPath)
        {
            //string rootDirectory = @"C:\Your\Root\Directory"; // Replace with your root directory
            TreeNode rootNode = new TreeNode(folderPath);
            tv.Nodes.Add(rootNode);
            PopulateTreeView(rootNode, folderPath);
        }
        /// <summary>
        /// 
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
        /// 
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
        public static TreeNode FindNodeByPath(TreeNodeCollection nodes, string path)
        {
            string[] nodeNames = path.Split('\\');
            TreeNode currentNode = null;
            foreach (string nodeName in nodeNames)
            {
                currentNode = nodes.Cast<TreeNode>().FirstOrDefault(n => n.Text == nodeName);
                if (currentNode == null)
                    break;
                nodes = currentNode.Nodes;
            }
            return currentNode;
        }

    }
}
