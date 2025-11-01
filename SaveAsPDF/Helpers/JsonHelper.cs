// Ignore Spelling: Json

using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Web.Script.Serialization;
using System.Windows.Forms;

namespace SaveAsPDF.Helpers
{
    internal class JsonTreeNode
    {
        public string Text { get; set; }
        public List<JsonTreeNode> Children { get; set; }
    }

    public static class JsonHelper
    {
        public static void SaveTreeToJson(TreeView treeView, string filePath)
        {
            if (treeView == null) throw new ArgumentNullException(nameof(treeView));
            if (string.IsNullOrEmpty(filePath)) throw new ArgumentException("filePath must be provided", nameof(filePath));

            var roots = new List<JsonTreeNode>();
            foreach (TreeNode node in treeView.Nodes)
            {
                roots.Add(ConvertToJsonNode(node));
            }

            var serializer = new JavaScriptSerializer { MaxJsonLength = Int32.MaxValue };
            string json = serializer.Serialize(roots);
            File.WriteAllText(filePath, json, Encoding.UTF8);
        }

        public static void LoadTreeFromJson(string filePath, TreeView treeView)
        {
            if (treeView == null) throw new ArgumentNullException(nameof(treeView));
            if (string.IsNullOrEmpty(filePath)) throw new ArgumentException("filePath must be provided", nameof(filePath));
            if (!File.Exists(filePath)) throw new FileNotFoundException("JSON tree file not found", filePath);

            string json = File.ReadAllText(filePath, Encoding.UTF8);
            var serializer = new JavaScriptSerializer { MaxJsonLength = Int32.MaxValue };
            var roots = serializer.Deserialize<List<JsonTreeNode>>(json);

            treeView.BeginUpdate();
            treeView.Nodes.Clear();
            try
            {
                if (roots != null)
                {
                    foreach (var rn in roots)
                    {
                        var tn = ConvertToTreeNode(rn);
                        treeView.Nodes.Add(tn);
                    }
                }
            }
            finally
            {
                treeView.EndUpdate();
            }
        }

        private static JsonTreeNode ConvertToJsonNode(TreeNode node)
        {
            var jn = new JsonTreeNode { Text = node.Text };
            if (node.Nodes != null && node.Nodes.Count > 0)
            {
                jn.Children = new List<JsonTreeNode>(node.Nodes.Count);
                foreach (TreeNode child in node.Nodes)
                {
                    jn.Children.Add(ConvertToJsonNode(child));
                }
            }
            return jn;
        }

        private static TreeNode ConvertToTreeNode(JsonTreeNode jn)
        {
            var tn = new TreeNode(jn.Text ?? string.Empty);
            if (jn.Children != null)
            {
                foreach (var child in jn.Children)
                {
                    tn.Nodes.Add(ConvertToTreeNode(child));
                }
            }
            return tn;
        }
    }
}
