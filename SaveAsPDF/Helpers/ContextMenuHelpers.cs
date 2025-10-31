using System;
using System.Diagnostics;
using System.IO;
using System.Windows.Forms;

namespace SaveAsPDF.Helpers
{
    /// <summary>
    /// Provides extension methods to enable context menus for RichTextBox, TextBox, and TreeView controls.
    /// </summary>
    public static class ContextMenuHelpers
    {
        private const string menuNameCut = "חתוך";
        private const string menuNameCopy = "העתק";
        private const string menuNamePaste = "הדבק";
        private const string menuNameSelectAll = "בחר הכל";
        private const string menuNameRefresh = "רענן";
        private const string menuNameDelete = "מחק";
        private const string menuNameUndo = "בטל";
        private const string menuNameRedo = "בצע שוב";
        private const string menuNameNew = "חדש";
        private const string menuNameRename = "שנה שם";
        private const string menuNameOpen = "פתח";

        /// <summary>
        /// Enables a context menu with common editing actions for a <see cref="RichTextBox"/>.
        /// </summary>
        /// <param name="rtxt">The <see cref="RichTextBox"/> to enable the context menu for.</param>
        public static void EnableContextMenu(this RichTextBox rtxt)
        {
            if (rtxt.ContextMenuStrip != null) return;

            var cms = new ContextMenuStrip { ShowImageMargin = false };

            var tsmiCut = new ToolStripMenuItem(menuNameCut, null, (s, e) => rtxt.Cut());
            var tsmiCopy = new ToolStripMenuItem(menuNameCopy, null, (s, e) => rtxt.Copy());
            var tsmiPaste = new ToolStripMenuItem(menuNamePaste, null, (s, e) => rtxt.Paste());
            var tsmiDelete = new ToolStripMenuItem(menuNameDelete, null, (s, e) => rtxt.SelectedText = "");
            var tsmiSelectAll = new ToolStripMenuItem(menuNameSelectAll, null, (s, e) => rtxt.SelectAll());
            var tsmiUndo = new ToolStripMenuItem(menuNameUndo, null, (s, e) => rtxt.Undo());
            var tsmiRedo = new ToolStripMenuItem(menuNameRedo, null, (s, e) => rtxt.Redo());

            cms.Items.AddRange(new ToolStripItem[] {
                tsmiCut, tsmiCopy, tsmiPaste, tsmiDelete,
                new ToolStripSeparator(),
                tsmiSelectAll,
                new ToolStripSeparator(),
                tsmiUndo, tsmiRedo
            });

            cms.Opening += (sender, e) =>
            {
                tsmiUndo.Enabled = !rtxt.ReadOnly && rtxt.CanUndo;
                tsmiRedo.Enabled = !rtxt.ReadOnly && rtxt.CanRedo;
                tsmiCut.Enabled = !rtxt.ReadOnly && rtxt.SelectionLength >0;
                tsmiCopy.Enabled = rtxt.SelectionLength >0;
                tsmiPaste.Enabled = !rtxt.ReadOnly && Clipboard.ContainsText();
                tsmiDelete.Enabled = !rtxt.ReadOnly && rtxt.SelectionLength >0;
                tsmiSelectAll.Enabled = rtxt.TextLength >0 && rtxt.SelectionLength < rtxt.TextLength;
            };

            rtxt.ContextMenuStrip = cms;
        }

        /// <summary>
        /// Enables a context menu with common editing actions for a <see cref="TextBox"/>.
        /// </summary>
        /// <param name="txt">The <see cref="TextBox"/> to enable the context menu for.</param>
        public static void EnableContextMenu(this TextBox txt)
        {
            if (txt.ContextMenuStrip != null) return;

            var cms = new ContextMenuStrip { ShowImageMargin = false };

            var tsmiCut = new ToolStripMenuItem(menuNameCut, null, (s, e) => txt.Cut());
            var tsmiCopy = new ToolStripMenuItem(menuNameCopy, null, (s, e) => txt.Copy());
            var tsmiPaste = new ToolStripMenuItem(menuNamePaste, null, (s, e) => txt.Paste());
            var tsmiDelete = new ToolStripMenuItem(menuNameDelete, null, (s, e) => txt.SelectedText = "");
            var tsmiSelectAll = new ToolStripMenuItem(menuNameSelectAll, null, (s, e) => txt.SelectAll());
            var tsmiUndo = new ToolStripMenuItem(menuNameUndo, null, (s, e) => txt.Undo());

            cms.Items.AddRange(new ToolStripItem[] {
                tsmiCut, tsmiCopy, tsmiPaste, tsmiDelete,
                new ToolStripSeparator(),
                tsmiSelectAll,
                new ToolStripSeparator(),
                tsmiUndo
            });

            cms.Opening += (sender, e) =>
            {
                tsmiUndo.Enabled = !txt.ReadOnly && txt.CanUndo;
                tsmiCut.Enabled = !txt.ReadOnly && txt.SelectionLength >0;
                tsmiCopy.Enabled = txt.SelectionLength >0;
                tsmiPaste.Enabled = !txt.ReadOnly && Clipboard.ContainsText();
                tsmiDelete.Enabled = !txt.ReadOnly && txt.SelectionLength >0;
                tsmiSelectAll.Enabled = txt.TextLength >0 && txt.SelectionLength < txt.TextLength;
            };

            txt.ContextMenuStrip = cms;
        }

        /// <summary>
        /// Enables a context menu for a <see cref="TreeView"/> with options for folder management.
        /// </summary>
        /// <param name="tv">The <see cref="TreeView"/> to enable the context menu for.</param>
        public static void EnableContextMenu(this TreeView tv)
        {
            if (tv.ContextMenuStrip != null) return;

            var cms = new ContextMenuStrip { ShowImageMargin = true };

            // New Folder
            var tsmiNew = new ToolStripMenuItem(menuNameNew);
            tsmiNew.Click += (sender, e) =>
            {
                try
                {
                    if (tv.SelectedNode == null)
                    {
                        XMessageBox.Show(
                            "יש לבחור תיקיה בעץ לפני יצירת תיקיה חדשה.",
                            "SaveAsPDF:EnableContextMenu",
                            XMessageBoxButtons.OK,
                            XMessageBoxIcon.Warning,
                            XMessageAlignment.Right,
                            XMessageLanguage.Hebrew
                        );
                        return;
                    }

                    string basePath = Path.Combine(FormMain.settingsModel.ProjectRootFolder.Parent.FullName, tv.SelectedNode.FullPath, "New Folder");
                    string[] tf = FileFoldersHelper.CreateDirectory(basePath).Split('\\');
                    tv.AddNode(tv.SelectedNode, tf[tf.Length -1]);
                }
                catch (Exception ex)
                {
                    XMessageBox.Show(
                        ex.Message,
                        "SaveAsPDF:EnableContextMenu",
                        XMessageBoxButtons.OK,
                        XMessageBoxIcon.Error,
                        XMessageAlignment.Right,
                        XMessageLanguage.Hebrew
                    );
                }
            };
            cms.Items.Add(tsmiNew);

            // Add Date Folder
            DateTime date = DateTime.Now;
            var tsmiAddDate = new ToolStripMenuItem(date.ToString("dd.MM.yyyy"));
            tsmiAddDate.Click += (sender, e) =>
            {
                try
                {
                    if (tv.SelectedNode == null)
                    {
                        XMessageBox.Show(
                            "יש לבחור תיקיה בעץ לפני יצירת תיקיה חדשה.",
                            "SaveAsPDF:EnableContextMenu",
                            XMessageBoxButtons.OK,
                            XMessageBoxIcon.Warning,
                            XMessageAlignment.Right,
                            XMessageLanguage.Hebrew
                        );
                        return;
                    }

                    string basePath = Path.Combine(FormMain.settingsModel.ProjectRootFolder.Parent.FullName, tv.SelectedNode.FullPath, date.ToString("dd.MM.yyyy"));
                    string[] tf = FileFoldersHelper.CreateDirectory(basePath).Split('\\');
                    tv.AddNode(tv.SelectedNode, tf[tf.Length -1]);
                }
                catch (Exception ex)
                {
                    XMessageBox.Show(
                        ex.Message,
                        "SaveAsPDF:EnableContextMenu",
                        XMessageBoxButtons.OK,
                        XMessageBoxIcon.Error,
                        XMessageAlignment.Right,
                        XMessageLanguage.Hebrew
                    );
                }
            };
            cms.Items.Add(tsmiAddDate);

            cms.Items.Add(new ToolStripSeparator());

            // Open Folder
            var tsmiOpen = new ToolStripMenuItem(menuNameOpen);
            tsmiOpen.Click += (sender, e) =>
            {
                if (tv.SelectedNode == null)
                {
                    XMessageBox.Show(
                        "יש לבחור תיקיה בעץ כדי לפתוח אותה.",
                        "SaveAsPDF:EnableContextMenu",
                        XMessageBoxButtons.OK,
                        XMessageBoxIcon.Warning,
                        XMessageAlignment.Right,
                        XMessageLanguage.Hebrew
                    );
                    return;
                }
                string path = Path.Combine(FormMain.settingsModel.ProjectRootFolder.Parent.FullName, tv.SelectedNode.FullPath);
                Process.Start(path);
            };
            cms.Items.Add(tsmiOpen);

            // Delete Folder
            var tsmiDelete = new ToolStripMenuItem(menuNameDelete);
            tsmiDelete.Click += (sender, e) =>
            {
                if (tv.SelectedNode?.Parent != null)
                {
                    string fullPath = Path.Combine(FormMain.settingsModel.ProjectRootFolder.Parent.FullName, tv.SelectedNode.FullPath);
                    var result = XMessageBox.Show(
                        "האם למחוק תיקייה ואת כל הקבצים והתיקיות שהיא מכילה?\n" +
                        fullPath,
                        "SaveAsPDF",
                        XMessageBoxButtons.YesNo,
                        XMessageBoxIcon.Warning,
                        XMessageAlignment.Right,
                        XMessageLanguage.Hebrew
                    );
                    if (result == DialogResult.Yes)
                    {
                        FileFoldersHelper.DeleteDirectory(fullPath);
                        TreeHelpers.DeleteNode(tv);
                    }
                }
                else
                {
                    XMessageBox.Show(
                        "לא ניתן למחוק את התיקייה הראשית בפרויקט",
                        "SaveAsPDF",
                        XMessageBoxButtons.OK,
                        XMessageBoxIcon.Warning,
                        XMessageAlignment.Right,
                        XMessageLanguage.Hebrew
                    );
                }
            };
            cms.Items.Add(tsmiDelete);

            // Rename Folder
            var tsmiRename = new ToolStripMenuItem(menuNameRename);
            tsmiRename.Click += (sender, e) =>
            {
                if (tv.SelectedNode == null)
                {
                    XMessageBox.Show(
                        "יש לבחור תיקיה בעץ כדי לשנות את שמה.",
                        "SaveAsPDF",
                        XMessageBoxButtons.OK,
                        XMessageBoxIcon.Warning,
                        XMessageAlignment.Right,
                        XMessageLanguage.Hebrew
                    );
                    return;
                }
                tv.RenameNode(tv.SelectedNode);
            };
            cms.Items.Add(tsmiRename);

            cms.Items.Add(new ToolStripSeparator());

            // Refresh
            var tsmiRefresh = new ToolStripMenuItem(menuNameRefresh);
            tsmiRefresh.Click += (sender, e) =>
            {
                tv.Nodes.Clear();
                tv.Nodes.Add(TreeHelpers.CreateDirectoryNode(FormMain.settingsModel.ProjectRootFolder));
                tv.ExpandAll();
                tv.SelectedNode = tv.Nodes[0];
            };
            cms.Items.Add(tsmiRefresh);

            cms.Items.Add(new ToolStripSeparator());

            cms.Opening += (sender, e) =>
            {
                tsmiRename.Enabled = tv.SelectedNode != null && tv.SelectedNode.Parent != null;
                tsmiDelete.Enabled = tv.SelectedNode != null && tv.SelectedNode.Parent != null;
            };

            tv.ContextMenuStrip = cms;
        }
    }
}