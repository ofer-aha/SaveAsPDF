// Ignore Spelling: // Ignore Spelling: frm מפשחה הכל יש לבחור הודעות דואר אלקטרוני בלבד אימייל הסר הכול שם קובץ גודל יש לבחור הודעות דואר אלקטרוני בלבד ההודעה נשמרה ב  תאריך  שמירה  שם הפרויקט  מס פרויקט  הערות  שם משתמש בחר  הסר  מספר פרויקט כפי שמופיע במסטרפלן שם לא חוקי  אין להשתמש בתווים הבאים  עריכת שם שם לא חוקי לא ניתן ליצור שם ריק חובה תו אחד לפחות עריכת שם מספר פרויקט לא חוקי
using SaveAsPDF.Helpers;
using SaveAsPDF.Models;
using System;
using System.IO;
using System.Windows.Controls;
using System.Windows.Forms;
using System.Xml;

namespace SaveAsPDF
{
    public partial class frmSettings : Form
    {
        private readonly ISettingsRequester callingForm;
        SettingsModel settingsModel = new SettingsModel();

        private TreeNode mySelectedNode;
        private bool _isDirty = false;


        public frmSettings(ISettingsRequester caller)
        {
            InitializeComponent();
            callingForm = caller;


            //List<string> fList = new List<string>();
            //foreach (TreeNode node in tvProjectSubFolders.Nodes)
            //{
            //    fList.AddRange(TreeHelper.ListNodesPath(node));
            //}
            //settingsModel.ProjectRootFolders = fList;
        }

        private void frmSettings_Load(object sender, EventArgs e)
        {
            //public frmSettings()

            //TODO1:frmSettings_Load
            // ON HOLD

            tvProjectSubFolders.Nodes.Add(frmMain.settingsModel.ProjectRootTag);
            tvProjectSubFolders.HideSelection = false;
            tvProjectSubFolders.PathSeparator = @"\";

            //load settings.settings to settingsModel 
            SettingsHelpers.loadSettingsToModel(settingsModel);

            txtRootFolder.Text = settingsModel.RootDrive;
            txtMinAttSize.Text = settingsModel.MinAttachmentSize.ToString();

            txtTreePath.Text = settingsModel.DefaultTreeFile;
            tvProjectSubFolders.LoadDefaultTree();
            tvProjectSubFolders.SelectedNode = tvProjectSubFolders.Nodes[0];

            cmbDefaultFolder.Items.Clear();

            TreeHelper.TvNodesToCombo(cmbDefaultFolder, tvProjectSubFolders.Nodes[0]);
            cmbDefaultFolder.SelectedIndex = settingsModel == null ? 0 : settingsModel.DefaultFolderID;

            //advanced settings tab 
            txtSaveAsPDFFolder.Text = settingsModel.XmlSaveAsPDFFolder;
            txtXmlProjectFile.Text = settingsModel.XmlProjectFile;
            txtXmlEmployeesFile.Text = settingsModel.XmlEmployeesFile;
            txtProjectRootTag.Text = settingsModel.ProjectRootTag;
            txtDateTag.Text = settingsModel.DateTag;

            //reset Dirty flag 
            _isDirty = false;
        }



        private void bntCancel_Click(object sender, EventArgs e)
        {
            Close();
        }

        /// <summary>
        /// Get the tree node under the mouse pointer and 
        /// save it in the _mySelectedNode variable.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void tvProjectSubFolders_MouseDown(object sender, MouseEventArgs e)
        {
            //TODO3: recheck - do i need it? 
            mySelectedNode = tvProjectSubFolders.GetNodeAt(e.X, e.Y);
        }

        private void tvProjectSubFolders_AfterLabelEdit(object sender, NodeLabelEditEventArgs e)
        {
            if (e.Label != null)
            {
                _isDirty = true;
                if (e.Label.Length > 0)
                {

                    if (e.Label.IndexOfAny(new char[] { '\\', '/', ':', '*', '?', '<', '>', '|', '"' }) == -1)
                    {
                        // Stop editing without canceling the label change.
                        e.Node.EndEdit(false);
                    }
                    else
                    {
                        /* Cancel the label edit action, inform the user, and
                           place the node in edit mode again. */
                        e.CancelEdit = true;
                        MessageBox.Show("שם לא חוקי.\n" +
                           "אין להשתמש בתווים הבאים '\\', '/', ':', '*', '?', '<', '>', '|', '\"' ",
                           "עריכת שם");
                        e.Node.BeginEdit();
                    }
                }
                else
                {
                    // Cancel the label edit action, inform the user, and
                    // place the node in edit mode again. 
                    e.CancelEdit = true;
                    MessageBox.Show("שם לא חוקי.\n לא ניתן ליצור שם ריק. חובה תו אחד לפחות", "עריכת שם");
                    e.Node.BeginEdit();
                }
            }
        }

        private void menueAdd_Click(object sender, EventArgs e)
        {
            tvProjectSubFolders.AddNode(mySelectedNode);
            _isDirty = true;
        }

        private void menuDel_Click(object sender, EventArgs e)
        {
            tvProjectSubFolders.DelNode(mySelectedNode);
            _isDirty = true;
        }

        private void menuRename_Click(object sender, EventArgs e)
        {
            tvProjectSubFolders.RenameNode(mySelectedNode);
            _isDirty = true;
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            if (_isDirty)
            {
                DialogResult result = MessageBox.Show("שמור שינויים?", "SaveAsPDF", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (result == DialogResult.Yes)
                {
                    SettingsHelpers.saveModelToSettings(settingsModel);
                }
            }
            //send the updated model back to MainForm 
            callingForm.SettingsComplete(settingsModel);
            _isDirty = false;
            Close();
        }

        private void btnSaveSettings_Click(object sender, EventArgs e)
        {
            SettingsHelpers.saveModelToSettings(settingsModel);
            _isDirty = false;
        }


        private void btnFolders_Click(object sender, EventArgs e)
        {
            var Dialog = new FolderPicker();

            Dialog.InputPath = settingsModel.RootDrive;

            if (Dialog.ShowDialog(Handle) == true)
            {
                _isDirty = true;
                txtRootFolder.Text = Dialog.ResultPath;
            }
        }

        private void btnLoadTreeFile_Click(object sender, EventArgs e)
        {
            OpenFileDialog dlg = new OpenFileDialog();
            dlg.Title = "פתח קובץ תיקיות";
            dlg.Filter = "קובץ תיקיות (*.fld)|*.fld";

            if (dlg.ShowDialog() == DialogResult.OK)
            {
                tvProjectSubFolders.LoadTreeViewFromFile(dlg.FileName);
                txtTreePath.Text = dlg.FileName;
                settingsModel.DefaultTreeFile = dlg.FileName;
                _isDirty = true;
            }

        }

        private void btnSaveAsTreeFile_Click(object sender, EventArgs ev)
        {
            try
            {
                SaveFileDialog dlg = new SaveFileDialog();
                dlg.Title = "שמור קובץ תיקיות";
                dlg.Filter = "קובץ תיקיות (*.fld)|*.fld";

                if (dlg.ShowDialog() == DialogResult.OK)
                {
                    dlg.CheckPathExists = true;
                    dlg.InitialDirectory = Path.GetDirectoryName(txtTreePath.Text);
                    TreeHelper.SaveTreeViewIntoFile(dlg.FileName, tvProjectSubFolders);
                    txtTreePath.Text = dlg.FileName;
                    settingsModel.DefaultTreeFile = dlg.FileName;
                    _isDirty = false;
                    //TODO3: NEXT VERSION: change the tree file to XML
                }

            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "frmSettings:btnSaveAsTreeFile_Click");
            }
        }
        private void btnSaveTreeFile_Click(object sender, EventArgs e)
        {
            if (File.Exists(txtTreePath.Text))
            {
                TreeHelper.SaveTreeViewIntoFile(txtTreePath.Text, tvProjectSubFolders);
            }
            else
            {

                DialogResult result = MessageBox.Show($"קובץ לא נימצא:\n{txtTreePath.Text}\n האם לשמור קובץ חדש?", "SaveAsPDF קובץ עץ סיפריות", MessageBoxButtons.YesNo);

                if (result == DialogResult.Yes)
                {
                    TreeHelper.SaveTreeViewIntoFile(txtTreePath.Text, tvProjectSubFolders);
                }
            }
        }

        private XmlNode GetXmlNode(TreeViewItem tnode, XmlDocument d)
        {
            XmlNode n = d.CreateNode(XmlNodeType.Element, tnode.Name, "");
            foreach (TreeViewItem t in tnode.Items)
            {
                n.AppendChild(GetXmlNode(t, d));
            }
            return n;
        }

        //TODO3: to check
        //private void btnDefaultTree_Click(object sender, System.EventArgs e)
        //{
        //    TreeHelper.RestTree(tvProjectSubFolders);
        //    settingsModel settingsModel = new settingsModel();
        //    List<string> fList = new List<string>();

        //    foreach (TreeNode node in tvProjectSubFolders.Nodes)
        //    {
        //        fList.AddRange(TreeHelper.ListNodesPath(node));
        //    }
        //    settingsModel.ProjectRootFolders = fList;

        //    cbDefaultFolder.DataSource = settingsModel.ProjectRootFolders;
        //    cbDefaultFolder.Refresh();

        //}


        private void btnLoadDefaultTree_Click(object sender, EventArgs e)
        {
            tvProjectSubFolders.LoadDefaultTree();
            _isDirty = true;
        }


        private void txtMinAttSize_TextChanged(object sender, EventArgs e)
        {
            _isDirty = true;
        }



        private void menuAddDate_Click(object sender, EventArgs e)
        {

            tvProjectSubFolders.AddNode(mySelectedNode, settingsModel.DateTag);
            _isDirty = true;

        }

        private void menuAppendDate_Click(object sender, EventArgs e)
        {
            //tvProjectSubFolders.SelectedNode.Name += settingsModel.dateTag;
            TreeHelper.RenameNode(tvProjectSubFolders, mySelectedNode, mySelectedNode.Name + settingsModel.DateTag);
        }

        private void cmbDefaultFolder_SelectedIndexChanged(object sender, EventArgs e)
        {
            //update settings 
            settingsModel.DefaultFolderID = (int)cmbDefaultFolder.SelectedIndex;
            settingsModel.ProjectRootFolders = new DirectoryInfo(cmbDefaultFolder.Text);
            _isDirty = true;
        }

        private void txtRootFolder_TextChanged(object sender, EventArgs e)
        {
            _isDirty = true;
        }

        private void txtSaveAsPDFFolder_TextChanged(object sender, EventArgs e)
        {
            _isDirty = true;
        }

        private void txtXmlProjectFile_TextChanged(object sender, EventArgs e)
        {
            _isDirty = true;
        }

        private void txtXmlEmployeesFile_TextChanged(object sender, EventArgs e)
        {
            _isDirty = true;
        }

        private void txtProjectRootTag_TextChanged(object sender, EventArgs e)
        {
            _isDirty = true;
        }

        private void txtDateTag_TextChanged(object sender, EventArgs e)
        {
            _isDirty = true;
        }

        private void txtTreePath_TextChanged(object sender, EventArgs e)
        {
            _isDirty = true;
        }

    }
}