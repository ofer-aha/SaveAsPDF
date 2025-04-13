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
    public partial class FormSettings : Form
    {
        private ISettingsRequester _callingForm;
        private SettingsModel _settingsModel = new SettingsModel();
        private TreeNode mySelectedNode;
        private bool _isDirty = false;


        public FormSettings(ISettingsRequester caller)
        {
            InitializeComponent();
            _callingForm = caller;

            _callingForm.SettingsComplete(_settingsModel);

            //load settings.settings to _settingsModel 
            //SettingsHelpers.LoadSettingsToModel(_settingsModel);

            tvProjectSubFolders.Nodes.Add(_settingsModel.ProjectRootTag);
            tvProjectSubFolders.HideSelection = false;
            tvProjectSubFolders.PathSeparator = @"\";

            lblSaveAsPDFFolder.Text = "תיקיית שמירת קבצי PDF";
            lblXmlProjectFile.Text = "קובץ פרויקטים";
            lblXmlEmployeesFile.Text = "קובץ עובדים";
            lblProjectRootTag.Text = "תגית שורש פרויקט";
            lblDateTag.Text = "תגית תאריך";
            lblTreePath.Text = "קובץ תיקיות";
            lblRootFolder.Text = "תיקיית ראשית";
            lblMinAttSize.Text = "גודל מינימלי לקובץ מצורף";

            txtRootFolder.Text = _settingsModel.RootDrive;
            txtMinAttSize.Text = _settingsModel.MinAttachmentSize.ToString();

            txtTreePath.Text = _settingsModel.DefaultTreeFile;

            try
            {
                string[] lines = File.ReadAllLines(_settingsModel.DefaultTreeFile);
                tvProjectSubFolders.LoadFromFile(lines);
                tvProjectSubFolders.ExpandAll();
                tvProjectSubFolders.SelectedNode = tvProjectSubFolders.Nodes[0];

            }
            catch (Exception e)
            {
                MessageBox.Show($"Error loading tree file: \n{e.Message}", "FormSettings");

            }

            cmbDefaultFolder.Items.Clear();
            TreeHelpers.ListNodesPath(tvProjectSubFolders.Nodes[0], cmbDefaultFolder);
            cmbDefaultFolder.SelectedIndex = _settingsModel == null ? 0 : _settingsModel.DefaultFolderID;


        }


        private void FormSettings_Activated(object sender, EventArgs e)
        {
            //refresh data if dirty? 

        }

        private void FormSettings_Load(object sender, EventArgs e)
        {

            //advanced settings tab 
            txtSaveAsPDFFolder.Text = _settingsModel.XmlSaveAsPDFFolder;
            txtXmlProjectFile.Text = _settingsModel.XmlProjectFile;
            txtXmlEmployeesFile.Text = _settingsModel.XmlEmployeesFile;
            txtProjectRootTag.Text = _settingsModel.ProjectRootTag;
            txtDateTag.Text = _settingsModel.DateTag;

            //Last projects
            lsbLastProjects.Items.Clear();
            foreach (var item in Settings.Default.LastProjects)
            {
                //don't add duplicates
                if (!lsbLastProjects.Items.Contains(item))
                {
                    lsbLastProjects.Items.Add(item);
                }

            }

            txtMaxProjectCount.Text = Settings.Default.MaxProjectCount.ToString();


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

        private void MenueAdd_Click(object sender, EventArgs e)
        {
            tvProjectSubFolders.AddNode(mySelectedNode);
            _isDirty = true;
        }

        private void MenuDel_Click(object sender, EventArgs e)
        {
            tvProjectSubFolders.DelNode(mySelectedNode);
            _isDirty = true;
        }

        private void MenuRename_Click(object sender, EventArgs e)
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
            openFileDialog.Title = "פתח קובץ תיקיות";
            openFileDialog.Filter = "קובץ תיקיות (*.fld)|*.fld";

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                string[] lines = File.ReadAllLines(openFileDialog.FileName);
                tvProjectSubFolders.LoadFromFile(lines);
                tvProjectSubFolders.ExpandAll();
                txtTreePath.Text = openFileDialog.FileName;
                _settingsModel.DefaultTreeFile = openFileDialog.FileName;
                cmbDefaultFolder.Items.Clear();
                TreeHelpers.ListNodesPath(tvProjectSubFolders.Nodes[0], cmbDefaultFolder);
                _isDirty = true;
            }

        }

        private void btnSaveTreeFile_Click(object sender, EventArgs e)
        {
            try
            {
                saveFileDialog.Title = "שמור קובץ תיקיות";
                saveFileDialog.Filter = "קובץ תיקיות (*.fld)|*.fld";

                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    saveFileDialog.CheckPathExists = true;
                    saveFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                    TreeHelpers.SaveTreeViewIntoFile(saveFileDialog.FileName, tvProjectSubFolders);
                    tvProjectSubFolders.ExpandAll();
                    txtTreePath.Text = saveFileDialog.FileName;
                    _settingsModel.DefaultTreeFile = saveFileDialog.FileName;
                    cmbDefaultFolder.Items.Clear();
                    TreeHelpers.ListNodesPath(tvProjectSubFolders.Nodes[0], cmbDefaultFolder);
                    _isDirty = false;
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "FormSettings: btnSaveTreeFile_Click");
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



        private void btnLoadDefaultTree_Click(object sender, EventArgs e)
        {
            //tvProjectSubFolders.LoadDefaultTree();
            tvProjectSubFolders.LoadFromList();
            cmbDefaultFolder.Items.Clear();
            TreeHelpers.ListNodesPath(tvProjectSubFolders.Nodes[0], cmbDefaultFolder);

            _isDirty = true;
        }


        private void txtMinAttSize_TextChanged(object sender, EventArgs e)
        {
            _isDirty = true;
        }



        private void MenuAddDate_Click(object sender, EventArgs e)
        {

            tvProjectSubFolders.AddNode(mySelectedNode, _settingsModel.DateTag);
            _isDirty = true;

        }

        private void MenuAppendDate_Click(object sender, EventArgs e)
        {
            //tvProjectSubFolders.SelectedNode.Name += _settingsModel.dateTag;
            TreeHelpers.RenameNode(tvProjectSubFolders, mySelectedNode, mySelectedNode.Name + _settingsModel.DateTag);
        }

        private void cmbDefaultFolder_SelectedIndexChanged(object sender, EventArgs e)
        {
            //update settings 
            //_settingsModel.DefaultFolderID = (int)cmbDefaultFolder.SelectedIndex;
            //_settingsModel.ProjectRootFolder = new DirectoryInfo(cmbDefaultFolder.Text);
            _settingsModel.DefaultSavePath = cmbDefaultFolder.Text;
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

        private void FormSettings_FormClosing(object sender, FormClosingEventArgs e)
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