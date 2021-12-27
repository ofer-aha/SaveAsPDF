using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;
using Microsoft.Office.Interop.Word;
using SaveAsPDF.Models;
using SaveAsPDF.Properties;

namespace SaveAsPDF.Helpers
{
    public  class TreeHelper
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
                    directoryNode.Nodes.Add(CreateDirectoryNode(directory));
                }


                //foreach (var file in directoryInfo.GetFiles())
                //{
                //    directoryNode.Nodes.Add(new TreeNode(file.Name) { ImageIndex = 0 });
                //}
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message, "SaveAsPDF");

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
                if (!IsHidden(d))
                {
                    DirectoryInfo di = new DirectoryInfo(d);
                    TreeNode tn = new TreeNode(di.Name);
                    tn.Tag = di;
                    int subCount = 0;
                    try { subCount = Directory.GetDirectories(d).Length; }
                    catch { /* ignore accessdenied */  }
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
        private static bool IsHidden (string sDir)
        {
            DirectoryInfo dir = new DirectoryInfo(sDir);
            if (dir.Attributes.HasFlag(FileAttributes.Hidden))
            {
                return true;
            }
            return false;
        }
        public static  TreeNode TraverseDirectory(string path)
        {
            TreeNode result = new TreeNode(path);
            Cursor.Current = Cursors.WaitCursor; //show the user we are doning somthing..... "the compuer is thinking"
            if (Directory.Exists(path))
            {
                foreach (string subdirectory in Directory.GetDirectories(path))
                {
                    if (!IsHidden(subdirectory))
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

        public static void RenameNode(TreeView treeView, TreeNode mySelectedNode)
        {
            if (mySelectedNode != null && mySelectedNode.Parent != null)
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

 
   

        public static void AddNode(TreeView treeView, TreeNode mySelectedNode)
        {
            mySelectedNode = treeView.SelectedNode.Nodes.Add("חדש");

            treeView.SelectedNode = mySelectedNode;
            mySelectedNode.Expand();

            RenameNode(treeView,mySelectedNode);
        }

        public static void DelNode(TreeView treeView, TreeNode mySelectedNode)
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
        /// Convert a node to List of Strings full path  
        /// </summary>
        /// <param name="oParentNode">The TreeNode</param>
        /// <returns>List of Strings </returns>
        //public static List<string> ListNodesPath(TreeNode oParentNode)
        //{
        //    List<String> output = new List<string>(); 

        //    // Start recursion on all subnodes.
        //    foreach (TreeNode oSubNode in oParentNode.Nodes)
        //    {
        //        if (oSubNode.Nodes.Count == 0)
        //        {
        //            output.Add(oSubNode.FullPath);
        //        }
        //        ListNodesPath(oSubNode);
        //    }
        //    return output;
        //}


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
        /// Write this node's subtree into the StringBuilder.
        /// </summary>
        /// <param name="level">Node Level as int.</param>
        /// <param name="node">Node object</param>
        /// <param name="sb">string line to be proccesed</param>
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
        /// Load a TreeView control from a file that uses tabs
        /// to show indentation.
        /// </summary>
        /// <param name="file_name"></param>
        /// <param name="trv"></param>
        public static void LoadTreeViewFromFile(string file_name, TreeView trv)
        {
            // Get the file's contents.
            string file_contents = File.ReadAllText(file_name);

            // Break the file into lines.
            string[] lines = file_contents.Split(
                new char[] { '\r', '\n' },
                StringSplitOptions.RemoveEmptyEntries);

            // Process the lines.
            trv.Nodes.Clear();
            Dictionary<int, TreeNode> parents =
                new Dictionary<int, TreeNode>();
            foreach (string text_line in lines)
            {
                // See how many tabs are at the start of the line.
                int level = text_line.Length -
                    text_line.TrimStart('\t').Length;

                // Add the new node.
                if (level == 0)
                    parents[level] = trv.Nodes.Add(text_line.Trim());
                else
                    parents[level] = parents[level - 1].Nodes.Add(text_line.Trim());
                parents[level].EnsureVisible();
            }
        }
        public static void RestTree(TreeView treeView)
        {
            treeView.Nodes.Clear();

            treeView.Nodes.Add("מספר_פרויקט");
            treeView.SelectedNode = treeView.Nodes[0];

            TreeNode node = treeView.SelectedNode.Nodes.Add("DWG");
            node.Parent.Nodes.Add("מכתבים");
            node.Parent.Nodes.Add("OLD");
            node.Parent.Nodes.Add("PDF");
            node = treeView.SelectedNode.Nodes.Add("אישור ציוד");
            node.Nodes.Add("לוח חשמל ראשי");
            node.Nodes.Add("דיזל גנרטור");
            node.Nodes.Add("גופי תאורה");
            node = treeView.SelectedNode.Nodes.Add("התקבל");
            node.Nodes.Add("אדריכלות");
            node.Nodes.Add("קונסטרוקציה");
            node.Nodes.Add("מיזוג אויר");
            node.Nodes.Add("בטיחות");
            node.Nodes.Add("כבישים");
            node = treeView.SelectedNode.Nodes.Add("נשלח");
            node.Nodes.Add("אדריכלות");
            node.Nodes.Add("קונסטרוקציה");
            node.Nodes.Add("מיזוג אויר");
            node.Nodes.Add("בטיחות");
            node.Nodes.Add("כבישים");

            treeView.ExpandAll();
        }
        /// <summary>
        /// Load the default folder tree from the default folder free file.
        /// If the folder tree file does not exist a default folder tree will be loaded
        /// </summary>
        /// <param name="treeView"></param>
        public static void LoadDefaultTree(TreeView treeView)
        {
            Assembly assembly = Assembly.GetExecutingAssembly();
            string location = assembly.CodeBase;
            string fullPath = new Uri(location).LocalPath; // path including the dll 
            string directoryPath = Path.GetDirectoryName(fullPath); // directory path 

            string defaultTreeFileName = directoryPath + "\\" + Settings.Default.defaultTreeFile;


            if (File.Exists(defaultTreeFileName))
            {
                TreeHelper.LoadTreeViewFromFile(defaultTreeFileName, treeView);
            }
            else
            {
                //Default tree file was not found -  create a new default free file
                string defaultTree = "<מספר פרויקט>\n" +
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
                TreeHelper.LoadTreeViewFromFile(defaultTreeFileName, treeView);

            }
        }


    }
}
