using SaveAsPDF.Helpers;
using SaveAsPDF.Models;
using SaveAsPDF.Properties;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace SaveAsPDF
{
    public partial class FormSettings : Form
    {
        private readonly ISettingsRequester _callingForm;
        private readonly SettingsModel _settingsModel = new SettingsModel();
        private TreeNode _mySelectedNode;
        private bool _isDirty;

        // Generic handlers to show short help text in the status strip
        private void Control_MouseEnterStatus(object sender, EventArgs e)
        {
            var ctl = sender as Control;
            toolStripStatusLabel.Text = ctl?.Tag as string ?? string.Empty;
        }

        private void Control_MouseLeaveStatus(object sender, EventArgs e)
        {
            toolStripStatusLabel.Text = string.Empty;
        }

        // Central wiring for hover help
        private void WireStatusHelp()
        {
            // Buttons
            btnOK.Tag = "אישור";
            btnOK.MouseEnter += Control_MouseEnterStatus; btnOK.MouseLeave += Control_MouseLeaveStatus;

            btnSaveSettings.Tag = "שמור שינויים";
            btnSaveSettings.MouseEnter += Control_MouseEnterStatus; btnSaveSettings.MouseLeave += Control_MouseLeaveStatus;

            bntCancel.Tag = "בטל וסגור";
            bntCancel.MouseEnter += Control_MouseEnterStatus; bntCancel.MouseLeave += Control_MouseLeaveStatus;

            btnFolders.Tag = "בחירת תיקיית שורש";
            btnFolders.MouseEnter += Control_MouseEnterStatus; btnFolders.MouseLeave += Control_MouseLeaveStatus;

            btnLoadTreeFile.Tag = "פתח קובץ מבנה";
            btnLoadTreeFile.MouseEnter += Control_MouseEnterStatus; btnLoadTreeFile.MouseLeave += Control_MouseLeaveStatus;

            btnSaveTreeFile.Tag = "שמור קובץ מבנה";
            btnSaveTreeFile.MouseEnter += Control_MouseEnterStatus; btnSaveTreeFile.MouseLeave += Control_MouseLeaveStatus;

            btnLoadDefaultTree.Tag = "טען מבנה ברירת מחדל";
            btnLoadDefaultTree.MouseEnter += Control_MouseEnterStatus; btnLoadDefaultTree.MouseLeave += Control_MouseLeaveStatus;

            btnClearList.Tag = "נקה רשימה";
            btnClearList.MouseEnter += Control_MouseEnterStatus; btnClearList.MouseLeave += Control_MouseLeaveStatus;

            // Inputs/labels
            txtRootFolder.Tag = "תיקיית שורש לפרויקטים";
            txtRootFolder.MouseEnter += Control_MouseEnterStatus; txtRootFolder.MouseLeave += Control_MouseLeaveStatus;

            cmbDefaultFolder.Tag = "בחר תיקיית ברירת מחדל";
            cmbDefaultFolder.MouseEnter += Control_MouseEnterStatus; cmbDefaultFolder.MouseLeave += Control_MouseLeaveStatus;

            txtMinAttSize.Tag = "סף גודל קובץ מצורף";
            txtMinAttSize.MouseEnter += Control_MouseEnterStatus; txtMinAttSize.MouseLeave += Control_MouseLeaveStatus;

            txtTreePath.Tag = "נתיב קובץ מבנה תיקיות";
            txtTreePath.MouseEnter += Control_MouseEnterStatus; txtTreePath.MouseLeave += Control_MouseLeaveStatus;

            tvProjectSubFolders.Tag = "עץ תיקיות";
            tvProjectSubFolders.MouseEnter += Control_MouseEnterStatus; tvProjectSubFolders.MouseLeave += Control_MouseLeaveStatus;

            txtSaveAsPDFFolder.Tag = "תיקיית SaveAsPDF";
            txtSaveAsPDFFolder.MouseEnter += Control_MouseEnterStatus; txtSaveAsPDFFolder.MouseLeave += Control_MouseLeaveStatus;

            txtXmlProjectFile.Tag = "קובץ XML פרויקט";
            txtXmlProjectFile.MouseEnter += Control_MouseEnterStatus; txtXmlProjectFile.MouseLeave += Control_MouseLeaveStatus;

            txtXmlEmployeesFile.Tag = "קובץ XML עובדים";
            txtXmlEmployeesFile.MouseEnter += Control_MouseEnterStatus; txtXmlEmployeesFile.MouseLeave += Control_MouseLeaveStatus;

            txtProjectRootTag.Tag = "תגית שורש פרויקט";
            txtProjectRootTag.MouseEnter += Control_MouseEnterStatus; txtProjectRootTag.MouseLeave += Control_MouseLeaveStatus;

            txtDateTag.Tag = "תגית תאריך";
            txtDateTag.MouseEnter += Control_MouseEnterStatus; txtDateTag.MouseLeave += Control_MouseLeaveStatus;

            lsbLastProjects.Tag = "פרויקטים אחרונים";
            lsbLastProjects.MouseEnter += Control_MouseEnterStatus; lsbLastProjects.MouseLeave += Control_MouseLeaveStatus;

            txtMaxProjectCount.Tag = "מס׳ פרויקטים אחרונים";
            txtMaxProjectCount.MouseEnter += Control_MouseEnterStatus; txtMaxProjectCount.MouseLeave += Control_MouseLeaveStatus;
        }

        public FormSettings(ISettingsRequester caller)
        {
            InitializeComponent();
            _callingForm = caller ?? throw new ArgumentNullException(nameof(caller));

            _callingForm.SettingsComplete(_settingsModel);

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

            LoadTreeFromFile(_settingsModel.DefaultTreeFile);

            cmbDefaultFolder.Items.Clear();
            PopulateComboBoxFromTree(tvProjectSubFolders, cmbDefaultFolder);
            
            // Find and select DefaultSavePath or use the saved DefaultFolderID
            
            // Case 1: If we have a DefaultSavePath, try to find a matching item
            if (!string.IsNullOrEmpty(_settingsModel.DefaultSavePath))
            {
                string pathToFind = _settingsModel.DefaultSavePath;
                bool found = false;
                
                // Try exact match first
                for (int i = 0; i < cmbDefaultFolder.Items.Count; i++)
                {
                    string itemText = cmbDefaultFolder.Items[i].ToString();
                    if (string.Equals(itemText, pathToFind, StringComparison.OrdinalIgnoreCase))
                    {
                        cmbDefaultFolder.SelectedIndex = i;
                        found = true;
                        break;
                    }
                }
                
                // If not found, try matching the last part of the path
                // This helps when DefaultSavePath contains project-specific parts
                if (!found && pathToFind.Contains("\\"))
                {
                    string[] pathParts = pathToFind.Split('\\');
                    string lastPart = pathParts.Length > 0 ? pathParts[pathParts.Length - 1] : string.Empty;
                    
                    if (!string.IsNullOrEmpty(lastPart))
                    {
                        for (int i = 0; i < cmbDefaultFolder.Items.Count; i++)
                        {
                            string itemText = cmbDefaultFolder.Items[i].ToString();
                            if (itemText.EndsWith("\\" + lastPart) || itemText == lastPart)
                            {
                                cmbDefaultFolder.SelectedIndex = i;
                                found = true;
                                break;
                            }
                        }
                    }
                }
                
                // If still not found and DefaultFolderID is valid, use it
                if (!found && cmbDefaultFolder.Items.Count > 0)
                {
                    if (_settingsModel.DefaultFolderID >= 0 && 
                        _settingsModel.DefaultFolderID < cmbDefaultFolder.Items.Count)
                    {
                        cmbDefaultFolder.SelectedIndex = _settingsModel.DefaultFolderID;
                    }
                    else
                    {
                        cmbDefaultFolder.SelectedIndex = 0;
                    }
                }
            }
            // Case 2: No DefaultSavePath, but we have a valid DefaultFolderID
            else if (cmbDefaultFolder.Items.Count > 0)
            {
                cmbDefaultFolder.SelectedIndex = (_settingsModel.DefaultFolderID >= 0 && 
                    _settingsModel.DefaultFolderID < cmbDefaultFolder.Items.Count)
                    ? _settingsModel.DefaultFolderID
                    : 0;
            }
        }

        private void LoadTreeFromFile(string filePath)
        {
            try
            {
                string[] lines = File.ReadAllLines(filePath);
                LoadTreeViewFromList(tvProjectSubFolders, lines);
                tvProjectSubFolders.ExpandAll();
                if (tvProjectSubFolders.Nodes.Count > 0)
                    tvProjectSubFolders.SelectedNode = tvProjectSubFolders.Nodes[0];
            }
            catch (Exception e)
            {
                XMessageBox.Show(
                    $"שגיאה בטעינת קובץ התיקיות: \n{e.Message}",
                    "FormSettings",
                    XMessageBoxButtons.OK,
                    XMessageBoxIcon.Error,
                    XMessageAlignment.Right,
                    XMessageLanguage.Hebrew
                );
            }
        }

        private void FormSettings_Load(object sender, EventArgs e)
        {
            txtSaveAsPDFFolder.Text = _settingsModel.XmlSaveAsPDFFolder;
            txtXmlProjectFile.Text = _settingsModel.XmlProjectFile;
            txtXmlEmployeesFile.Text = _settingsModel.XmlEmployeesFile;
            txtProjectRootTag.Text = _settingsModel.ProjectRootTag;
            txtDateTag.Text = _settingsModel.DateTag;

            lsbLastProjects.Items.Clear();
            foreach (string item in Settings.Default.LastProjects)
            {
                if (!lsbLastProjects.Items.Contains(item))
                {
                    lsbLastProjects.Items.Add(item);
                }
            }

            txtMaxProjectCount.Text = Settings.Default.MaxProjectCount.ToString();
            _isDirty = false;

            // Wire status help hover
            WireStatusHelp();
        }

        private void bntCancel_Click(object sender, EventArgs e) => Close();

        private void tvProjectSubFolders_MouseDown(object sender, MouseEventArgs e)
        {
            _mySelectedNode = tvProjectSubFolders.GetNodeAt(e.X, e.Y);
        }

        private void tvProjectSubFolders_AfterLabelEdit(object sender, NodeLabelEditEventArgs e)
        {
            if (e.Label == null || string.IsNullOrEmpty(e.Label) || e.Label.IndexOfAny(new char[] { '\\', '/', ':', '*', '?', '<', '>', '|', '"' }) != -1)
            {
                e.CancelEdit = true;
                XMessageBox.Show(
                    e.Label == null || string.IsNullOrEmpty(e.Label)
                        ? "שם לא חוקי.\n לא ניתן ליצור שם ריק. חובה תו אחד לפחות"
                        : "שם לא חוקי.\nאין להשתמש בתווים הבאים '\\', '/', ':', '*', '?', '<', '>', '|', '\"' ",
                    "עריכת שם",
                    XMessageBoxButtons.OK,
                    XMessageBoxIcon.Warning,
                    XMessageAlignment.Right,
                    XMessageLanguage.Hebrew
                );
                e.Node?.BeginEdit();
                return;
            }

            _isDirty = true;
            e.Node.EndEdit(false);
        }

        private void MenueAdd_Click(object sender, EventArgs e)
        {
            TreeHelpers.AddNode(tvProjectSubFolders, _mySelectedNode);
            _isDirty = true;
            
            // Refresh comboBox after adding a node
            cmbDefaultFolder.Items.Clear();
            PopulateComboBoxFromTree(tvProjectSubFolders, cmbDefaultFolder);
        }

        private void MenuDel_Click(object sender, EventArgs e)
        {
            TreeHelpers.DeleteNode(tvProjectSubFolders);
            _isDirty = true;
            
            // Refresh comboBox after deleting a node
            cmbDefaultFolder.Items.Clear();
            PopulateComboBoxFromTree(tvProjectSubFolders, cmbDefaultFolder);
        }

        private void MenuRename_Click(object sender, EventArgs e)
        {
            TreeHelpers.RenameNode(tvProjectSubFolders, _mySelectedNode);
            _isDirty = true;
            
            // Refresh comboBox after renaming a node
            cmbDefaultFolder.Items.Clear();
            PopulateComboBoxFromTree(tvProjectSubFolders, cmbDefaultFolder);
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            if (_isDirty)
            {
                DialogResult result = XMessageBox.Show(
                    this,
                    "שמור שינויים?",
                    "SaveAsPDF",
                    XMessageBoxButtons.YesNo,
                    XMessageBoxIcon.Question,
                    XMessageAlignment.Right,
                    XMessageLanguage.Hebrew
                );
                if (result == DialogResult.Yes)
                {
                    // Make sure to update DefaultSavePath before saving
                    if (cmbDefaultFolder.SelectedItem != null)
                    {
                        string selectedPath = cmbDefaultFolder.SelectedItem.ToString();

                        if (_settingsModel.ProjectRootTag != null && selectedPath.Contains(_settingsModel.ProjectRootTag))
                        {
                            _settingsModel.DefaultSavePath = selectedPath;
                        }
                        else
                        {
                            _settingsModel.DefaultSavePath = selectedPath;
                        }

                        _settingsModel.DefaultFolderID = cmbDefaultFolder.SelectedIndex;
                    }

                    SettingsHelpers.SaveModelToSettings(_settingsModel);
                }
            }
            _callingForm.SettingsComplete(_settingsModel);
            _isDirty = false;
            Close();
        }

        private void btnSaveSettings_Click(object sender, EventArgs e)
        {
            // Make sure to update DefaultSavePath before saving
            if (cmbDefaultFolder.SelectedItem != null)
            {
                string selectedPath = cmbDefaultFolder.SelectedItem.ToString();
                
                // Check if the path contains the project root tag
                if (_settingsModel.ProjectRootTag != null && 
                    selectedPath.Contains(_settingsModel.ProjectRootTag))
                {
                    // This is a path with a project tag placeholder - store as is
                    _settingsModel.DefaultSavePath = selectedPath;
                }
                else
                {
                    // Regular path - store as is
                    _settingsModel.DefaultSavePath = selectedPath;
                }
                
                _settingsModel.DefaultFolderID = cmbDefaultFolder.SelectedIndex;
            }
            
            SettingsHelpers.SaveModelToSettings(_settingsModel);
            _isDirty = false;
        }

        private void btnFolders_Click(object sender, EventArgs e)
        {
            var dialog = new FolderPicker { InputPath = _settingsModel.RootDrive };
            if (dialog.ShowDialog(Handle) == true)
            {
                _isDirty = true;
                txtRootFolder.Text = dialog.ResultPath;
            }
        }
        
        private void btnLoadTreeFile_Click(object sender, EventArgs e)
        {
            openFileDialog.Title = "פתח קובץ תיקיות";
            openFileDialog.Filter = "קובץ תיקיות (*.fld)|*.fld";

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                LoadTreeFromFile(openFileDialog.FileName);
                txtTreePath.Text = openFileDialog.FileName;
                _settingsModel.DefaultTreeFile = openFileDialog.FileName;
                cmbDefaultFolder.Items.Clear();
                PopulateComboBoxFromTree(tvProjectSubFolders, cmbDefaultFolder);
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
                    SaveTreeViewToFile(tvProjectSubFolders, saveFileDialog.FileName);
                    tvProjectSubFolders.ExpandAll();
                    txtTreePath.Text = saveFileDialog.FileName;
                    _settingsModel.DefaultTreeFile = saveFileDialog.FileName;
                    cmbDefaultFolder.Items.Clear();
                    PopulateComboBoxFromTree(tvProjectSubFolders, cmbDefaultFolder);
                    _isDirty = false;
                }
            }
            catch (Exception ex)
            {
                XMessageBox.Show(
                    ex.Message,
                    "FormSettings: btnSaveTreeFile_Click",
                    XMessageBoxButtons.OK,
                    XMessageBoxIcon.Error,
                    XMessageAlignment.Right,
                    XMessageLanguage.Hebrew
                );
            }
        }

        private void LoadTreeViewFromList(TreeView treeView, string[] lines)
        {
            treeView.Nodes.Clear();
            Dictionary<int, TreeNode> indentToNodeDict = new Dictionary<int, TreeNode>();
            
            foreach (string line in lines)
            {
                if (string.IsNullOrWhiteSpace(line))
                    continue;
                
                // Determine indentation level (number of spaces at start)
                int indentLevel = 0;
                for (int i = 0; i < line.Length; i++)
                {
                    if (line[i] == ' ')
                        indentLevel++;
                    else
                        break;
                }
                
                // Get the node text without leading spaces
                string nodeText = line.Trim();
                
                // Create the node
                TreeNode newNode = new TreeNode(nodeText);
                
                if (indentLevel == 0)
                {
                    // Root node
                    treeView.Nodes.Add(newNode);
                    indentToNodeDict[indentLevel] = newNode;
                }
                else
                {
                    // Find the parent node based on indentation
                    int parentIndentLevel = -1;
                    for (int i = indentLevel - 1; i >= 0; i--)
                    {
                        if (indentToNodeDict.ContainsKey(i))
                        {
                            parentIndentLevel = i;
                            break;
                        }
                    }
                    
                    if (parentIndentLevel >= 0)
                    {
                        indentToNodeDict[parentIndentLevel].Nodes.Add(newNode);
                        indentToNodeDict[indentLevel] = newNode;
                    }
                    else
                    {
                        // If we can't find a parent, add as a root node
                        treeView.Nodes.Add(newNode);
                        indentToNodeDict[indentLevel] = newNode;
                    }
                }
            }
        }

        private void SaveTreeViewToFile(TreeView treeView, string fileName)
        {
            using (var writer = new StreamWriter(fileName, false))
            {
                foreach (TreeNode node in treeView.Nodes)
                {
                    WriteNodeRecursive(writer, node, "");
                }
            }
        }

        private void WriteNodeRecursive(StreamWriter writer, TreeNode node, string indent)
        {
            writer.WriteLine($"{indent}{node.Text}");
            foreach (TreeNode child in node.Nodes)
            {
                WriteNodeRecursive(writer, child, indent + "  ");
            }
        }

        private void PopulateComboBoxFromTree(TreeView treeView, ComboBox comboBox)
        {
            comboBox.Items.Clear();
            foreach (TreeNode node in treeView.Nodes)
            {
                AddNodePathsToComboBox(node, comboBox);
            }
        }

        private void AddNodePathsToComboBox(TreeNode node, ComboBox comboBox)
        {
            // Only add the node path, not the combined project path
            // This ensures we're storing just the logical structure,
            // not a physical path that could cause duplication later
            string nodePath = node.FullPath;
            
            // Don't include any project ID placeholder in the path
            // This prevents duplicating the project ID in the final path
            if (!string.IsNullOrEmpty(nodePath) && 
                !comboBox.Items.Contains(nodePath))
            {
                comboBox.Items.Add(nodePath);
            }
            
            // Process child nodes
            foreach (TreeNode child in node.Nodes)
            {
                AddNodePathsToComboBox(child, comboBox);
            }
        }

        private void btnLoadDefaultTree_Click(object sender, EventArgs e)
        {
            tvProjectSubFolders.Nodes.Clear();
            tvProjectSubFolders.Nodes.Add(_settingsModel.ProjectRootTag);
            cmbDefaultFolder.Items.Clear();
            PopulateComboBoxFromTree(tvProjectSubFolders, cmbDefaultFolder);
            _isDirty = true;
        }

        private void txtMinAttSize_TextChanged(object sender, EventArgs e)
        {
            _isDirty = true;
        }

        private void MenuAddDate_Click(object sender, EventArgs e)
        {
            TreeHelpers.AddNode(tvProjectSubFolders, _mySelectedNode, _settingsModel.DateTag);
            _isDirty = true;
            
            // Refresh comboBox after adding a date node
            cmbDefaultFolder.Items.Clear();
            PopulateComboBoxFromTree(tvProjectSubFolders, cmbDefaultFolder);
        }

        private void MenuAppendDate_Click(object sender, EventArgs e)
        {
            if (_mySelectedNode != null)
            {
                _mySelectedNode.Text += _settingsModel.DateTag;
                _isDirty = true;
                
                // Refresh comboBox after modifying a node
                cmbDefaultFolder.Items.Clear();
                PopulateComboBoxFromTree(tvProjectSubFolders, cmbDefaultFolder);
            }
        }

        private void cmbDefaultFolder_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cmbDefaultFolder.SelectedItem != null)
            {
                // Store the relative path without any project ID information
                string selectedPath = cmbDefaultFolder.SelectedItem.ToString();

                // If contains project tag, keep as-is; otherwise store as-is
                if (_settingsModel.ProjectRootTag != null && selectedPath.Contains(_settingsModel.ProjectRootTag))
                {
                    _settingsModel.DefaultSavePath = selectedPath;
                }
                else
                {
                    _settingsModel.DefaultSavePath = selectedPath;
                }

                _settingsModel.DefaultFolderID = cmbDefaultFolder.SelectedIndex;
                _isDirty = true;

                // Persist immediately so FormMain can pick it up on next project selection
                SettingsHelpers.SaveModelToSettings(_settingsModel);

                // Optionally notify caller about updated settings
                _callingForm?.SettingsComplete(_settingsModel);
            }
        }

        private void txtRootFolder_TextChanged(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(txtRootFolder.Text))
            {
                _settingsModel.RootDrive = txtRootFolder.Text;
                _isDirty = true;
            }
        }

        private void txtSaveAsPDFFolder_TextChanged(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(txtSaveAsPDFFolder.Text))
            {
                _settingsModel.XmlSaveAsPDFFolder = txtSaveAsPDFFolder.Text;
                _isDirty = true;
            }
        }

        private void txtXmlProjectFile_TextChanged(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(txtXmlProjectFile.Text))
            {
                _settingsModel.XmlProjectFile = txtXmlProjectFile.Text;
                _isDirty = true;
            }
        }

        private void txtXmlEmployeesFile_TextChanged(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(txtXmlEmployeesFile.Text))
            {
                _settingsModel.XmlEmployeesFile = txtXmlEmployeesFile.Text;
                _isDirty = true;
            }
        }

        private void txtProjectRootTag_TextChanged(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(txtProjectRootTag.Text))
            {
                _settingsModel.ProjectRootTag = txtProjectRootTag.Text;
                _isDirty = true;
            }
        }

        private void txtDateTag_TextChanged(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(txtDateTag.Text))
            {
                _settingsModel.DateTag = txtDateTag.Text;
                _isDirty = true;
            }
        }

        private void txtTreePath_TextChanged(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(txtTreePath.Text))
            {
                _settingsModel.DefaultTreeFile = txtTreePath.Text;
                _isDirty = true;
            }
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
            if (int.TryParse(txtMaxProjectCount.Text, out int count))
            {
                Settings.Default.MaxProjectCount = count;
                Settings.Default.Save();
                errorProviderSettings.Clear();
                toolStripStatusLabel.Text = errorProviderSettings.GetError(txtMaxProjectCount);
                _isDirty = true;
            }
        }

        private void txtMaxProjectCount_Validating(object sender, System.ComponentModel.CancelEventArgs e)
        {
            if (!int.TryParse(txtMaxProjectCount.Text, out int count) || count < 0 || count > 99)
            {
                e.Cancel = true;
                errorProviderSettings.SetError(txtMaxProjectCount, "ניתן להזין ערכים בין 0 ל- 99. ערך מומלץ 10");
                txtMaxProjectCount.Select(0, txtMaxProjectCount.Text.Length);
                toolStripStatusLabel.Text = errorProviderSettings.GetError(txtMaxProjectCount);
            }
        }

        private void txtMinAttSize_Validated(object sender, EventArgs e)
        {
            if (int.TryParse(txtMinAttSize.Text, out int count))
            {
                _settingsModel.MinAttachmentSize = count;
                errorProviderSettings.Clear();
                toolStripStatusLabel.Text = errorProviderSettings.GetError(txtMinAttSize);
                _isDirty = true;
            }
        }

        private void txtMinAttSize_Validating(object sender, System.ComponentModel.CancelEventArgs e)
        {
            if (!int.TryParse(txtMinAttSize.Text, out int count) || count < 0 || count > 9999)
            {
                e.Cancel = true;
                errorProviderSettings.SetError(txtMinAttSize, "ניתן להזין ערכים בין 0 ל- 9999. ערך מומלץ 8192");
                txtMinAttSize.Select(0, txtMinAttSize.Text.Length);
                toolStripStatusLabel.Text = errorProviderSettings.GetError(txtMinAttSize);
            }
        }
    }
}