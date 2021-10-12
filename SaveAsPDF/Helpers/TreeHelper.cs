using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using SaveAsPDF.Models;

namespace SaveAsPDF.Helpers
{
    public static class TreeHelper
    {
        public static TreeNode CreateDirectoryNode(DirectoryInfo directoryInfo)
        {
            var directoryNode = new TreeNode(directoryInfo.Name) { ImageIndex = 0 };

            foreach (var directory in directoryInfo.GetDirectories())
            {
                directoryNode.Nodes.Add(CreateDirectoryNode(directory));
            }


            foreach (var file in directoryInfo.GetFiles())
            {
                directoryNode.Nodes.Add(new TreeNode(file.Name) { ImageIndex = 1 });
            }
            return directoryNode;
        }

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
    }
}
