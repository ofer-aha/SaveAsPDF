// Ignore Spelling: // Ignore Spelling: frm מפשחה הכל יש לבחור הודעות דואר אלקטרוני בלבד אימייל הסר הכול שם קובץ גודל יש לבחור הודעות דואר אלקטרוני בלבד ההודעה נשמרה ב  תאריך  שמירה  שם הפרויקט  מס פרויקט  הערות  שם משתמש בחר  הסר  מספר פרויקט כפי שמופיע במסטרפלן שם לא חוקי  אין להשתמש בתווים הבאים  עריכת שם שם לא חוקי לא ניתן ליצור שם ריק חובה תו אחד לפחות עריכת שם מספר פרויקט לא חוקי
using SaveAsPDF.Helpers;
using SaveAsPDF.Models;
using SaveAsPDF.Properties;
using System;
using System.IO;
using System.Windows.Controls;
using System.Windows.Forms;
using System.Xml;

namespace SaveAsPDF
{
    public partial class frmSettings : Form
    {
        private ISettingsRequester _callingForm;
        private SettingsModel _settingsModel = new SettingsModel();
        private TreeNode mySelectedNode;
        private bool _isDirty = false;


        public frmSettings(ISettingsRequester caller)
        {
            InitializeComponent();
            _callingForm = caller;


            //List<string> fList = new List<string>();
            //foreach (TreeNode node in tvProjectSubFolders.Nodes)
            //{
            //    fList.AddRange(TreeHelper.ListNodesPath(node));
            //}
            //_settingsModel.ProjectRootFolder = fList;


            _callingForm.SettingsComplete(_settingsModel);

            //load settings.settings to _settingsModel 
            //SettingsHelpers.LoadSettingsToModel(_settingsModel);

            tvProjectSubFolders.Nodes.Add(_settingsModel.ProjectRootTag);
            tvProjectSubFolders.HideSelection = false;
            tvProjectSubFolders.PathSeparator = @"\";


            txtRootFolder.Text = _settingsModel.RootDrive;
            txtMinAttSize.Text = _settingsModel.MinAttachmentSize.ToString();

            txtTreePath.Text = _settingsModel.DefaultTreeFile;
            tvProjectSubFolders.LoadDefaultTree();
            tvProjectSubFolders.SelectedNode = tvProjectSubFolders.Nodes[0];

            cmbDefaultFolder.Items.Clear();

            TreeHelpers.TvNodesToCombo(cmbDefaultFolder, tvProjectSubFolders.Nodes[0]);
            cmbDefaultFolder.SelectedIndex = _settingsModel == null ? 0 : _settingsModel.DefaultFolderID;

            //advanced settings tab 
            txtSaveAsPDFFolder.Text = _settingsModel.XmlSaveAsPDFFolder;
            txtXmlProjectFile.Text = _settingsModel.XmlProjectFile;
            txtXmlEmployeesFile.Text = _settingsModel.XmlEmployeesFile;
            txtProjectRootTag.Text = _settingsModel.ProjectRootTag;
            txtDateTag.Text = _settingsModel.DateTag;


        }


        private void frmSettings_Activated(object sender, EventArgs e)
        {
            //refresh data if dirty? 

        }

        private void frmSettings_Load(object sender, EventArgs e)
        {
            //public frmSettings()

            //TODO1:frmSettings_Load
            // ON HOLD

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
                    SettingsHelpers.SaveModelToSettings(_settingsModel);
                }
            }
            //send the updated model back to MainForm 
            _callingForm.SettingsComplete(_settingsModel);
            _isDirty = false;
            Close();
        }

        private void btnSaveSettings_Click(object sender, EventArgs e)
        {
            SettingsHelpers.SaveModelToSettings(_settingsModel);
            _isDirty = false;
        }


        private void btnFolders_Click(object sender, EventArgs e)
        {
            var Dialog = new FolderPicker();

            Dialog.InputPath = _settingsModel.RootDrive;

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
                _settingsModel.DefaultTreeFile = dlg.FileName;
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
                    TreeHelpers.SaveTreeViewIntoFile(dlg.FileName, tvProjectSubFolders);
                    txtTreePath.Text = dlg.FileName;
                    _settingsModel.DefaultTreeFile = dlg.FileName;
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
                TreeHelpers.SaveTreeViewIntoFile(txtTreePath.Text, tvProjectSubFolders);
            }
            else
            {

                DialogResult result = MessageBox.Show($"קובץ לא נימצא:\n{txtTreePath.Text}\n האם לשמור קובץ חדש?", "SaveAsPDF קובץ עץ סיפריות", MessageBoxButtons.YesNo);

                if (result == DialogResult.Yes)
                {
                    TreeHelpers.SaveTreeViewIntoFile(txtTreePath.Text, tvProjectSubFolders);
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
        //    _settingsModel _settingsModel = new _settingsModel();
        //    List<string> fList = new List<string>();

        //    foreach (TreeNode node in tvProjectSubFolders.Nodes)
        //    {
        //        fList.AddRange(TreeHelper.ListNodesPath(node));
        //    }
        //    _settingsModel.ProjectRootFolder = fList;

        //    cbDefaultFolder.DataSource = _settingsModel.ProjectRootFolder;
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

            tvProjectSubFolders.AddNode(mySelectedNode, _settingsModel.DateTag);
            _isDirty = true;

        }

        private void menuAppendDate_Click(object sender, EventArgs e)
        {
            //tvProjectSubFolders.SelectedNode.Name += _settingsModel.dateTag;
            TreeHelpers.RenameNode(tvProjectSubFolders, mySelectedNode, mySelectedNode.Name + _settingsModel.DateTag);
        }

        private void cmbDefaultFolder_SelectedIndexChanged(object sender, EventArgs e)
        {
            //update settings 
            _settingsModel.DefaultFolderID = (int)cmbDefaultFolder.SelectedIndex;
            _settingsModel.ProjectRootFolder = new DirectoryInfo(cmbDefaultFolder.Text);
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

        private void frmSettings_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (_isDirty)
            {
                btnOK_Click(sender, e);
            }
        }

        private void txtMaxProjectCount_Validated(object sender, EventArgs e)
        {
            int.TryParse(txtMaxProjectCount.Text, out int count);
            Settings.Default.MaxProjectCount = count;
            Settings.Default.Save();
            errorProviderSettings.Clear();
            toolStripStatusLabel.Text = errorProviderSettings.GetError(txtMaxProjectCount);
            _isDirty = true;

        }

        private void txtMaxProjectCount_Validating(object sender, System.ComponentModel.CancelEventArgs e)
        {
            if (int.TryParse(txtMaxProjectCount.Text, out int count))
            {
                if (count < 0 || count > 99)
                {
                    e.Cancel = true;
                    errorProviderSettings.SetError(txtMaxProjectCount, "ניתן להזין ערכים בין 0 ל- 99. ערך מומלץ 10");
                    txtMaxProjectCount.Select(0, txtMaxProjectCount.Text.Length);
                    toolStripStatusLabel.Text = errorProviderSettings.GetError(txtMaxProjectCount);
                }
            }
            else
            {
                e.Cancel = true;
                errorProviderSettings.SetError(txtMaxProjectCount, "יש להזין ספרות בלבד");
                txtMaxProjectCount.Select(0, txtMaxProjectCount.Text.Length);
                toolStripStatusLabel.Text = errorProviderSettings.GetError(txtMaxProjectCount);
            }
        }

        private void txtMinAttSize_Validating(object sender, System.ComponentModel.CancelEventArgs e)
        {
            if (int.TryParse(txtMinAttSize.Text, out int count))
            {
                if (count < 0 || count > 9999)
                {
                    e.Cancel = true;
                    errorProviderSettings.SetError(txtMinAttSize, "ניתן להזין ערכים בין 0 ל- 9999. ערך מומלץ 8192");
                    txtMinAttSize.Select(0, txtMinAttSize.Text.Length);
                    toolStripStatusLabel.Text = errorProviderSettings.GetError(txtMinAttSize);
                }
            }
            else
            {
                e.Cancel = true;
                errorProviderSettings.SetError(txtMinAttSize, "יש להזין ספרות בלבד");
                txtMinAttSize.Select(0, txtMinAttSize.Text.Length);
                toolStripStatusLabel.Text = errorProviderSettings.GetError(txtMinAttSize);
            }

        }

        private void txtMinAttSize_Validated(object sender, EventArgs e)
        {
            errorProviderSettings.Clear();
            toolStripStatusLabel.Text = errorProviderSettings.GetError(txtMinAttSize);
            _isDirty = true;
        }
    }
}