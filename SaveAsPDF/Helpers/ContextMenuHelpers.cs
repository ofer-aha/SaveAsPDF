// Ignore Spelling: rtxt txt

using System;
using System.Diagnostics;
using System.IO;
using System.Windows.Forms;
namespace SaveAsPDF.Helpers
{

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
        /// Enable the context menu to the RichTextxtox 
        /// </summary>
        /// <param name="rtxt">RichTextxtox object</param>
        public static void EnableContextMenu(this RichTextBox rtxt)
        {
            if (rtxt.ContextMenuStrip == null)
            {
                // Create a ContextMenuStrip without icons
                ContextMenuStrip cms = new ContextMenuStrip();
                cms.ShowImageMargin = false;

                // 3. Add the Cut option (cuts the selected text inside the richTextxBox)
                ToolStripMenuItem tsmiCut = new ToolStripMenuItem(menuNameCut);
                tsmiCut.Click += (sender, e) => rtxt.Cut();
                cms.Items.Add(tsmiCut);

                // 4. Add the Copy option (copies the selected text inside the richTextxBox)
                ToolStripMenuItem tsmiCopy = new ToolStripMenuItem(menuNameCopy);
                tsmiCopy.Click += (sender, e) => rtxt.Copy();
                cms.Items.Add(tsmiCopy);

                // 5. Add the Paste option (adds the text from the clipboard into the richTextxBox)
                ToolStripMenuItem tsmiPaste = new ToolStripMenuItem(menuNamePaste);
                tsmiPaste.Click += (sender, e) => rtxt.Paste();
                cms.Items.Add(tsmiPaste);

                // 6. Add the Delete Option (remove the selected text in the richTextxBox)
                ToolStripMenuItem tsmiDelete = new ToolStripMenuItem(menuNameDelete);
                tsmiDelete.Click += (sender, e) => rtxt.SelectedText = "";
                cms.Items.Add(tsmiDelete);

                // Add a Separator
                cms.Items.Add(new ToolStripSeparator());

                // 7. Add the Select All Option (selects all the text inside the richTextxBox)
                ToolStripMenuItem tsmiSelectAll = new ToolStripMenuItem(menuNameSelectAll);
                tsmiSelectAll.Click += (sender, e) => rtxt.SelectAll();
                cms.Items.Add(tsmiSelectAll);

                // Add a Separator
                cms.Items.Add(new ToolStripSeparator());

                // 1. Add the Undo option
                ToolStripMenuItem tsmiUndo = new ToolStripMenuItem(menuNameUndo);
                tsmiUndo.Click += (sender, e) => rtxt.Undo();
                cms.Items.Add(tsmiUndo);

                // 2. Add the Redo option
                ToolStripMenuItem tsmiRedo = new ToolStripMenuItem(menuNameRedo);
                tsmiRedo.Click += (sender, e) => rtxt.Redo();
                cms.Items.Add(tsmiRedo);

                // When opening the menu, check if the condition is fulfilled 
                // in order to enable the action
                cms.Opening += (sender, e) =>
                {
                    tsmiUndo.Enabled = !rtxt.ReadOnly && rtxt.CanUndo;
                    tsmiRedo.Enabled = !rtxt.ReadOnly && rtxt.CanRedo;
                    tsmiCut.Enabled = !rtxt.ReadOnly && rtxt.SelectionLength > 0;
                    tsmiCopy.Enabled = rtxt.SelectionLength > 0;
                    tsmiPaste.Enabled = !rtxt.ReadOnly && Clipboard.ContainsText();
                    tsmiDelete.Enabled = !rtxt.ReadOnly && rtxt.SelectionLength > 0;
                    tsmiSelectAll.Enabled = rtxt.TextLength > 0 && rtxt.SelectionLength < rtxt.TextLength;
                };

                rtxt.ContextMenuStrip = cms;
            }
        }
        /// <summary>
        /// Enable the context menu to the TextxBox 
        /// </summary>
        /// <param name="txt">TextxBox object</param>
        public static void EnableContextMenu(this TextBox txt)
        {
            if (txt.ContextMenuStrip == null)
            {
                // Create a ContextMenuStrip without icons
                ContextMenuStrip cms = new ContextMenuStrip();
                cms.ShowImageMargin = false;

                // 3. Add the Cut option (cuts the selected text inside the richTextxBox)
                ToolStripMenuItem tsmiCut = new ToolStripMenuItem(menuNameCut);
                tsmiCut.Click += (sender, e) => txt.Cut();
                cms.Items.Add(tsmiCut);

                // 4. Add the Copy option (copies the selected text inside the richTextxBox)
                ToolStripMenuItem tsmiCopy = new ToolStripMenuItem(menuNameCopy);
                tsmiCopy.Click += (sender, e) => txt.Copy();
                cms.Items.Add(tsmiCopy);

                // 5. Add the Paste option (adds the text from the clipboard into the richTextxBox)
                ToolStripMenuItem tsmiPaste = new ToolStripMenuItem(menuNamePaste);
                tsmiPaste.Click += (sender, e) => txt.Paste();
                cms.Items.Add(tsmiPaste);

                // 6. Add the Delete Option (remove the selected text in the richTextxBox)
                ToolStripMenuItem tsmiDelete = new ToolStripMenuItem(menuNameDelete);
                tsmiDelete.Click += (sender, e) => txt.SelectedText = "";
                cms.Items.Add(tsmiDelete);

                // Add a Separator
                cms.Items.Add(new ToolStripSeparator());

                // 7. Add the Select All Option (selects all the text inside the richTextxBox)
                ToolStripMenuItem tsmiSelectAll = new ToolStripMenuItem(menuNameSelectAll);
                tsmiSelectAll.Click += (sender, e) => txt.SelectAll();
                cms.Items.Add(tsmiSelectAll);

                // Add a Separator
                cms.Items.Add(new ToolStripSeparator());

                // 1. Add the Undo option
                ToolStripMenuItem tsmiUndo = new ToolStripMenuItem(menuNameUndo);
                tsmiUndo.Click += (sender, e) => txt.Undo();
                cms.Items.Add(tsmiUndo);

                // When opening the menu, check if the condition is fulfilled 
                // in order to enable the action
                cms.Opening += (sender, e) =>
                {
                    tsmiUndo.Enabled = !txt.ReadOnly && txt.CanUndo;
                    tsmiCut.Enabled = !txt.ReadOnly && txt.SelectionLength > 0;
                    tsmiCopy.Enabled = txt.SelectionLength > 0;
                    tsmiPaste.Enabled = !txt.ReadOnly && Clipboard.ContainsText();
                    tsmiDelete.Enabled = !txt.ReadOnly && txt.SelectionLength > 0;
                    tsmiSelectAll.Enabled = txt.TextLength > 0 && txt.SelectionLength < txt.TextLength;
                };

                txt.ContextMenuStrip = cms;
            }
        }


        /// <summary>
        /// Enable the context menu to the Tree-view 
        /// </summary>
        /// <param name="tv">TreeView object</param>
        public static void EnableContextMenu(this TreeView tv)
        {
            if (tv.ContextMenuStrip == null)
            {
                // Create a ContextMenuStrip without icons
                ContextMenuStrip cms = new ContextMenuStrip();
                cms.ShowImageMargin = true;

                // Add the New option (Create a new folder and new tree-node)
                ToolStripMenuItem tsmiNew = new ToolStripMenuItem(menuNameNew);
                tsmiNew.Click += (sender, e) =>
                {
                    try
                    {
                        string[] tf = FileFoldersHelper.CreateDirectory($@"{FormMain.settingsModel.ProjectRootFolder.Parent.FullName}\{FormMain.mySelectedNode.FullPath}\New Folder").Split('\\');
                        tv.AddNode(FormMain.mySelectedNode, tf[tf.Length - 1]);

                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message, "SaveAsPDF:EnableContextMenu");
                    }
                };

                cms.Items.Add(tsmiNew);

                // Add the tsmiAddDate option (Add a folder name with today's date)
                DateTime date = DateTime.Now;
                ToolStripMenuItem tsmiAddDate = new ToolStripMenuItem(date.ToString("dd.MM.yyyy"));
                tsmiAddDate.Click += (sender, e) =>
                {
                    try
                    {
                        string[] tf = FileFoldersHelper.CreateDirectory($@"{FormMain.settingsModel.ProjectRootFolder.Parent.FullName}\{FormMain.mySelectedNode.FullPath}\{date.ToString("dd.MM.yyyy")}").Split('\\');
                        tv.AddNode(FormMain.mySelectedNode, tf[tf.Length - 1]);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message, "SaveAsPDF:EnableContextMenu");
                    }

                };
                cms.Items.Add(tsmiAddDate);

                // Add a Separator
                cms.Items.Add(new ToolStripSeparator());


                //Add the Open option (adds the text from the clipboard into the richTextxBox)
                ToolStripMenuItem tsmiOpen = new ToolStripMenuItem(menuNameOpen);
                tsmiOpen.Click += (sender, e) =>
                {
                    //TODO: maybe this will work with open PDF 

                    Process.Start($@"{FormMain.settingsModel.ProjectRootFolder.Parent.FullName}\{FormMain.mySelectedNode.FullPath}");
                };
                cms.Items.Add(tsmiOpen);

                // Add the Delete Option (remove the selected folder and tree-node)
                ToolStripMenuItem tsmiDelete = new ToolStripMenuItem(menuNameDelete);
                tsmiDelete.Click += (sender, e) =>
                {
                    if (FormMain.mySelectedNode.Parent != null)
                    {
                        if (MessageBox.Show("האם למחוק תיקייה ואת כל הקבצים והתיקיות שהיא מכילה?\n" +
                                $@"{FormMain.settingsModel.ProjectRootFolder.Parent.FullName}\{FormMain.mySelectedNode.FullPath}", "SaveAsPDF", MessageBoxButtons.YesNo) == DialogResult.Yes)
                        {
                            FileFoldersHelper.DeleteDirectory($@"{FormMain.settingsModel.ProjectRootFolder.Parent.FullName}\{FormMain.mySelectedNode.FullPath}");
                            tv.DelNode(FormMain.mySelectedNode);
                        }
                    }
                    else
                    {
                        MessageBox.Show("לא ניתן למחוק את התיקייה הראשית בפרויקט", "SaveAsPDF");
                    }
                };
                cms.Items.Add(tsmiDelete);

                //Add the Rename option (Rename the folder and tree-node
                ToolStripMenuItem tsmiRename = new ToolStripMenuItem(menuNameRename);
                tsmiRename.Click += (sender, e) =>
                {
                    string oldName = $@"{FormMain.settingsModel.ProjectRootFolder.Parent.FullName}\frmMain.{FormMain.mySelectedNode.FullPath}";
                    DirectoryInfo directoryInfo = new DirectoryInfo(oldName.SafeFolderName());

                    tv.RenameNode(FormMain.mySelectedNode);
                    //tvFolders.Refresh();
                    FormMain.mySelectedNode = tv.SelectedNode;

                };
                cms.Items.Add(tsmiRename);

                // Add a Separator
                cms.Items.Add(new ToolStripSeparator());

                //Add the Refresh Option (refresh folder state)
                ToolStripMenuItem tsmiRefresh = new ToolStripMenuItem(menuNameRefresh);
                tsmiRefresh.Click += (sender, e) =>
                {
                    tv.Nodes.Clear();
                    tv.Nodes.Add(TreeHelpers.CreateDirectoryNode(FormMain.settingsModel.ProjectRootFolder));
                    tv.ExpandAll();
                    tv.SelectedNode = tv.Nodes[0];
                };
                cms.Items.Add(tsmiRefresh);

                // Add a Separator
                cms.Items.Add(new ToolStripSeparator());


                // When opening the menu, check if the condition is fulfilled 
                // in order to enable the action
                cms.Opening += (sender, e) =>
                {
                    tsmiRename.Enabled = tv.SelectedNode.Parent != null;
                    tsmiDelete.Enabled = tv.SelectedNode.Parent != null; //Cant delete the root node
                    //tsmiSelectAll.Enabled = tv.TextLength > 0 && tv.SelectionLength < tv.TextLength;
                };
                tv.ContextMenuStrip = cms;
            }
        }
    }
}