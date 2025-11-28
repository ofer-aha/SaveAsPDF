using SaveAsPDF.Helpers;
using SaveAsPDF.Models;
using SaveAsPDF.Properties;
using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;
using System.Drawing;

namespace SaveAsPDF
{
    public partial class FormSettings : Form
    {
        private readonly ISettingsRequester _callingForm;
        private readonly SettingsModel _settings_model = new SettingsModel();
        private TreeNode _mySelectedNode;
        private bool _isDirty;

        // Generic handlers to show short help text in the status strip
        private void ControlMouseEnterStatus(object sender, EventArgs e)
        {
            var ctl = sender as Control;
            toolStripStatusLabel.Text = ctl?.Tag as string ?? string.Empty;
        }

        private void ControlMouseLeaveStatus(object sender, EventArgs e)
        {
            toolStripStatusLabel.Text = string.Empty;
        }

        // Central wiring for hover help
        private void WireStatusHelp()
        {
            // Buttons
            btnOK.Tag = "אישור";
            btnOK.MouseEnter += ControlMouseEnterStatus;
            btnOK.MouseLeave += ControlMouseLeaveStatus;

            btnSaveSettings.Tag = "שמור שינויים";
            btnSaveSettings.MouseEnter += ControlMouseEnterStatus;
            btnSaveSettings.MouseLeave += ControlMouseLeaveStatus;

            bntCancel.Tag = "בטל וסגור";
            bntCancel.MouseEnter += ControlMouseEnterStatus;
            bntCancel.MouseLeave += ControlMouseLeaveStatus;

            btnFolders.Tag = "בחירת תיקיית שורש";
            btnFolders.MouseEnter += ControlMouseEnterStatus;
            btnFolders.MouseLeave += ControlMouseLeaveStatus;

            btnLoadTreeFile.Tag = "פתח קובץ מבנה (JSON)";
            btnLoadTreeFile.MouseEnter += ControlMouseEnterStatus;
            btnLoadTreeFile.MouseLeave += ControlMouseLeaveStatus;

            btnSaveTreeFile.Tag = "שמור קובץ מבנה (JSON)";
            btnSaveTreeFile.MouseEnter += ControlMouseEnterStatus;
            btnSaveTreeFile.MouseLeave += ControlMouseLeaveStatus;

            btnLoadDefaultTree.Tag = "טען מבנה ברירת מחדל";
            btnLoadDefaultTree.MouseEnter += ControlMouseEnterStatus;
            btnLoadDefaultTree.MouseLeave += ControlMouseLeaveStatus;

            btnClearList.Tag = "נקה רשימה";
            btnClearList.MouseEnter += ControlMouseEnterStatus;
            btnClearList.MouseLeave += ControlMouseLeaveStatus;

            // Inputs/labels
            txtRootFolder.Tag = "תיקיית שורש לפרויקטים";
            txtRootFolder.MouseEnter += ControlMouseEnterStatus;
            txtRootFolder.MouseLeave += ControlMouseLeaveStatus;

            cmbDefaultFolder.Tag = "בחר תיקיית ברירת מחדל";
            cmbDefaultFolder.MouseEnter += ControlMouseEnterStatus;
            cmbDefaultFolder.MouseLeave += ControlMouseLeaveStatus;

            txtMinAttSize.Tag = "סף גודל קובץ מצורף";
            txtMinAttSize.MouseEnter += ControlMouseEnterStatus;
            txtMinAttSize.MouseLeave += ControlMouseLeaveStatus;

            txtTreePath.Tag = "נתיב קובץ מבנה תיקיות";
            txtTreePath.MouseEnter += ControlMouseEnterStatus;
            txtTreePath.MouseLeave += ControlMouseLeaveStatus;

            tvProjectSubFolders.Tag = "עץ תיקיות";
            tvProjectSubFolders.MouseEnter += ControlMouseEnterStatus;
            tvProjectSubFolders.MouseLeave += ControlMouseLeaveStatus;

            txtSaveAsPDFFolder.Tag = "תיקיית SaveAsPDF";
            txtSaveAsPDFFolder.MouseEnter += ControlMouseEnterStatus;
            txtSaveAsPDFFolder.MouseLeave += ControlMouseLeaveStatus;

            txtXmlProjectFile.Tag = "קובץ XML פרויקט";
            txtXmlProjectFile.MouseEnter += ControlMouseEnterStatus;
            txtXmlProjectFile.MouseLeave += ControlMouseLeaveStatus;

            txtXmlEmployeesFile.Tag = "קובץ XML עובדים";
            txtXmlEmployeesFile.MouseEnter += ControlMouseEnterStatus;
            txtXmlEmployeesFile.MouseLeave += ControlMouseLeaveStatus;

            txtProjectRootTag.Tag = "תגית שורש פרויקט";
            txtProjectRootTag.MouseEnter += ControlMouseEnterStatus;
            txtProjectRootTag.MouseLeave += ControlMouseLeaveStatus;

            txtDateTag.Tag = "תגית תאריך";
            txtDateTag.MouseEnter += ControlMouseEnterStatus;
            txtDateTag.MouseLeave += ControlMouseLeaveStatus;

            lsbLastProjects.Tag = "פרויקטים אחרונים";
            lsbLastProjects.MouseEnter += ControlMouseEnterStatus;
            lsbLastProjects.MouseLeave += ControlMouseLeaveStatus;

            txtMaxProjectCount.Tag = "מס׳ פרויקטים אחרונים";
            txtMaxProjectCount.MouseEnter += ControlMouseEnterStatus;
            txtMaxProjectCount.MouseLeave += ControlMouseLeaveStatus;
        }

        public FormSettings(ISettingsRequester caller)
        {
            InitializeComponent();
            _callingForm = caller ?? throw new ArgumentNullException(nameof(caller));

            _callingForm.SettingsComplete(_settings_model);

            tvProjectSubFolders.HideSelection = false;
            tvProjectSubFolders.PathSeparator = "\\";

            // Enable drag & drop to allow moving nodes between folders
            tvProjectSubFolders.AllowDrop = true;
            tvProjectSubFolders.ItemDrag += TvProjectSubFolders_ItemDrag;
            tvProjectSubFolders.DragEnter += TvProjectSubFolders_DragEnter;
            tvProjectSubFolders.DragOver += TvProjectSubFolders_DragOver;
            tvProjectSubFolders.DragDrop += TvProjectSubFolders_DragDrop;

            lblSaveAsPDFFolder.Text = "תיקיית שמירת קבצי PDF";
            lblXmlProjectFile.Text = "קובץ פרויקטים";
            lblXmlEmployeesFile.Text = "קובץ עובדים";
            lblProjectRootTag.Text = "תגית שורש פרויקט";
            lblDateTag.Text = "תגית תאריך";
            lblTreePath.Text = "קובץ תיקיות";
            lblRootFolder.Text = "תיקיית ראשית";
            lblMinAttSize.Text = "גודל מינימלי לקובץ מצורף";

            txtRootFolder.Text = _settings_model.RootDrive;
            txtMinAttSize.Text = _settings_model.MinAttachmentSize.ToString();
            txtTreePath.Text = _settings_model.DefaultTreeFile;

            // Load default tree only if JSON file exists
            if (!string.IsNullOrEmpty(_settings_model.DefaultTreeFile) && File.Exists(_settings_model.DefaultTreeFile) && _settings_model.DefaultTreeFile.EndsWith(".json", StringComparison.OrdinalIgnoreCase))
            {
                try
                {
                    JsonHelper.LoadTreeFromJson(_settings_model.DefaultTreeFile, tvProjectSubFolders);
                }
                catch
                {
                    // ignore load errors here; user may load manually
                }
            }

            cmbDefaultFolder.Items.Clear();
            PopulateComboBoxFromTree(tvProjectSubFolders, cmbDefaultFolder);

            // Find and select DefaultSavePath or use the saved DefaultFolderID

            // Case1: If we have a DefaultSavePath, try to find a matching item
            if (!string.IsNullOrEmpty(_settings_model.DefaultSavePath))
            {
                string pathToFind = _settings_model.DefaultSavePath;
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
                    if (_settings_model.DefaultFolderID >= 0 &&
                        _settings_model.DefaultFolderID < cmbDefaultFolder.Items.Count)
                    {
                        cmbDefaultFolder.SelectedIndex = _settings_model.DefaultFolderID;
                    }
                    else
                    {
                        cmbDefaultFolder.SelectedIndex = 0;
                    }
                }
            }
            // Case2: No DefaultSavePath, but we have a valid DefaultFolderID
            else if (cmbDefaultFolder.Items.Count > 0)
            {
                cmbDefaultFolder.SelectedIndex = (_settings_model.DefaultFolderID >= 0 &&
                    _settings_model.DefaultFolderID < cmbDefaultFolder.Items.Count)
                    ? _settings_model.DefaultFolderID
                    : 0;
            }
        }

        private void FormSettings_Load(object sender, EventArgs e)
        {
            txtSaveAsPDFFolder.Text = _settings_model.XmlSaveAsPDFFolder;
            txtXmlProjectFile.Text = _settings_model.XmlProjectFile;
            txtXmlEmployeesFile.Text = _settings_model.XmlEmployeesFile;
            txtProjectRootTag.Text = _settings_model.ProjectRootTag;
            txtDateTag.Text = _settings_model.DateTag;

            lsbLastProjects.Items.Clear();
            foreach (string item in Settings.Default.LastProjects)
            {
                if (!lsbLastProjects.Items.Contains(item))
                {
                    lsbLastProjects.Items.Add(item);
                }
            }

            // Display the current number of last-projects in the textbox
            txtMaxProjectCount.Text = lsbLastProjects.Items.Count.ToString();
            _isDirty = false;

            // Wire status help hover
            WireStatusHelp();
        }

        private void bntCancel_Click(object sender, EventArgs e)
        {
            // Treat explicit cancel as discard: avoid save prompt on FormClosing
            _isDirty = false;
            DialogResult = DialogResult.Cancel;
            Close();
        }

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

                        if (_settings_model.ProjectRootTag != null && selectedPath.Contains(_settings_model.ProjectRootTag))
                        {
                            _settings_model.DefaultSavePath = selectedPath;
                        }
                        else
                        {
                            _settings_model.DefaultSavePath = selectedPath;
                        }

                        _settings_model.DefaultFolderID = cmbDefaultFolder.SelectedIndex;
                    }

                    SettingsHelpers.SaveModelToSettings(_settings_model);
                }
            }
            _callingForm.SettingsComplete(_settings_model);
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
                if (_settings_model.ProjectRootTag != null &&
                    selectedPath.Contains(_settings_model.ProjectRootTag))
                {
                    // This is a path with a project tag placeholder - store as is
                    _settings_model.DefaultSavePath = selectedPath;
                }
                else
                {
                    // Regular path - store as is
                    _settings_model.DefaultSavePath = selectedPath;
                }

                _settings_model.DefaultFolderID = cmbDefaultFolder.SelectedIndex;
            }

            SettingsHelpers.SaveModelToSettings(_settings_model);
            _isDirty = false;
        }

        private void btnFolders_Click(object sender, EventArgs e)
        {
            var dialog = new FolderPicker { InputPath = _settings_model.RootDrive };
            if (dialog.ShowDialog(Handle) == true)
            {
                _isDirty = true;
                txtRootFolder.Text = dialog.ResultPath;
            }
        }

        private void btnLoadTreeFile_Click(object sender, EventArgs e)
        {
            openFileDialog.Title = "פתח קובץ מבנה (JSON)";
            openFileDialog.Filter = "JSON tree (*.json)|*.json";

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                var file = openFileDialog.FileName;
                try
                {
                    SaveAsPDF.Helpers.JsonHelper.LoadTreeFromJson(file, tvProjectSubFolders);
                    txtTreePath.Text = file;
                    _settings_model.DefaultTreeFile = file;
                    cmbDefaultFolder.Items.Clear();
                    PopulateComboBoxFromTree(tvProjectSubFolders, cmbDefaultFolder);
                    _isDirty = true;
                }
                catch (Exception ex)
                {
                    XMessageBox.Show($"שגיאה בטעינת קובץ JSON: {ex.Message}", "FormSettings", XMessageBoxButtons.OK, XMessageBoxIcon.Error, XMessageAlignment.Right, XMessageLanguage.Hebrew);
                }
            }
        }

        private void btnSaveTreeFile_Click(object sender, EventArgs e)
        {
            try
            {
                saveFileDialog.Title = "שמור קובץ מבנה (JSON)";
                saveFileDialog.Filter = "JSON tree (*.json)|*.json";

                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    var file = saveFileDialog.FileName;
                    saveFileDialog.CheckPathExists = true;
                    saveFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);

                    // Save as JSON using helper
                    SaveAsPDF.Helpers.JsonHelper.SaveTreeToJson(tvProjectSubFolders, file);

                    tvProjectSubFolders.ExpandAll();
                    txtTreePath.Text = file;
                    _settings_model.DefaultTreeFile = file;
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

        private void PopulateComboBoxFromTree(TreeView treeView, ComboBox comboBox)
        {
            comboBox.Items.Clear();
            if (treeView == null) return;

            // Ensure PathSeparator is set so TreeNode.FullPath produces expected strings
            if (string.IsNullOrEmpty(treeView.PathSeparator))
                treeView.PathSeparator = "\\";

            // Traverse all nodes iteratively to avoid deep recursion
            var stack = new Stack<TreeNode>();
            // Push roots in reverse so processing order matches the visual order
            for (int r = treeView.Nodes.Count - 1; r >= 0; r--)
            {
                stack.Push(treeView.Nodes[r]);
            }

            while (stack.Count > 0)
            {
                var node = stack.Pop();
                if (node == null) continue;

                string fullPath = node.FullPath ?? string.Empty;
                if (!string.IsNullOrEmpty(fullPath))
                {
                    comboBox.Items.Add(fullPath);
                }

                // Push children so they are processed (in-order: push in reverse to preserve order)
                if (node.Nodes != null && node.Nodes.Count > 0)
                {
                    for (int i = node.Nodes.Count - 1; i >= 0; i--)
                    {
                        stack.Push(node.Nodes[i]);
                    }
                }
            }
        }

        private void btnLoadDefaultTree_Click(object sender, EventArgs e)
        {
            tvProjectSubFolders.Nodes.Clear();
            tvProjectSubFolders.Nodes.Add(_settings_model.ProjectRootTag);
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
            TreeHelpers.AddNode(tvProjectSubFolders, _mySelectedNode, _settings_model.DateTag);
            _isDirty = true;

            // Refresh comboBox after adding a date node
            cmbDefaultFolder.Items.Clear();
            PopulateComboBoxFromTree(tvProjectSubFolders, cmbDefaultFolder);
        }

        private void MenuAppendDate_Click(object sender, EventArgs e)
        {
            if (_mySelectedNode != null)
            {
                _mySelectedNode.Text += _settings_model.DateTag;
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
                if (_settings_model.ProjectRootTag != null && selectedPath.Contains(_settings_model.ProjectRootTag))
                {
                    _settings_model.DefaultSavePath = selectedPath;
                }
                else
                {
                    _settings_model.DefaultSavePath = selectedPath;
                }

                _settings_model.DefaultFolderID = cmbDefaultFolder.SelectedIndex;
                _isDirty = true;

                // Persist immediately so FormMain can pick it up on next project selection
                SettingsHelpers.SaveModelToSettings(_settings_model);

                // Optionally notify caller about updated settings
                _callingForm?.SettingsComplete(_settings_model);
            }
        }

        private void txtRootFolder_TextChanged(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(txtRootFolder.Text))
            {
                _settings_model.RootDrive = txtRootFolder.Text;
                _isDirty = true;
            }
        }

        private void txtSaveAsPDFFolder_TextChanged(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(txtSaveAsPDFFolder.Text))
            {
                _settings_model.XmlSaveAsPDFFolder = txtSaveAsPDFFolder.Text;
                _isDirty = true;
            }
        }

        private void txtXmlProjectFile_TextChanged(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(txtXmlProjectFile.Text))
            {
                _settings_model.XmlProjectFile = txtXmlProjectFile.Text;
                _isDirty = true;
            }
        }

        private void txtXmlEmployeesFile_TextChanged(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(txtXmlEmployeesFile.Text))
            {
                _settings_model.XmlEmployeesFile = txtXmlEmployeesFile.Text;
                _isDirty = true;
            }
        }

        private void txtProjectRootTag_TextChanged(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(txtProjectRootTag.Text))
            {
                _settings_model.ProjectRootTag = txtProjectRootTag.Text;
                _isDirty = true;
            }
        }

        private void txtDateTag_TextChanged(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(txtDateTag.Text))
            {
                _settings_model.DateTag = txtDateTag.Text;
                _isDirty = true;
            }
        }

        private void txtTreePath_TextChanged(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(txtTreePath.Text))
            {
                _settings_model.DefaultTreeFile = txtTreePath.Text;
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
                _settings_model.MinAttachmentSize = count;
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

        // Drag & Drop handlers for moving TreeNodes
        private void TvProjectSubFolders_ItemDrag(object sender, ItemDragEventArgs e)
        {
            if (e.Item is TreeNode node)
            {
                DoDragDrop(node, DragDropEffects.Move);
            }
        }

        private void TvProjectSubFolders_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(typeof(TreeNode)))
                e.Effect = DragDropEffects.Move;
            else
                e.Effect = DragDropEffects.None;
        }

        private void TvProjectSubFolders_DragOver(object sender, DragEventArgs e)
        {
            var tv = sender as TreeView;
            if (tv == null) return;

            if (!e.Data.GetDataPresent(typeof(TreeNode)))
            {
                e.Effect = DragDropEffects.None;
                return;
            }

            Point clientPoint = tv.PointToClient(new Point(e.X, e.Y));
            TreeNode targetNode = tv.GetNodeAt(clientPoint);

            // Highlight the node under cursor
            tv.SelectedNode = targetNode;
            e.Effect = DragDropEffects.Move;
        }

        private void TvProjectSubFolders_DragDrop(object sender, DragEventArgs e)
        {
            var tv = sender as TreeView;
            if (tv == null) return;

            if (!e.Data.GetDataPresent(typeof(TreeNode))) return;

            TreeNode draggedNode = e.Data.GetData(typeof(TreeNode)) as TreeNode;
            if (draggedNode == null) return;

            Point clientPoint = tv.PointToClient(new Point(e.X, e.Y));
            TreeNode targetNode = tv.GetNodeAt(clientPoint);

            // Prevent dropping a node onto itself or its descendant
            if (targetNode != null)
            {
                TreeNode cur = targetNode;
                while (cur != null)
                {
                    if (cur == draggedNode) return; // invalid
                    cur = cur.Parent;
                }
            }

            tv.BeginUpdate();
            try
            {
                // Remove from original location
                if (draggedNode.Parent != null)
                    draggedNode.Parent.Nodes.Remove(draggedNode);
                else
                    tv.Nodes.Remove(draggedNode);

                // Add to new location
                if (targetNode == null)
                    tv.Nodes.Add(draggedNode);
                else
                    targetNode.Nodes.Add(draggedNode);

                // Expand and select moved node
                draggedNode.EnsureVisible();
                tv.SelectedNode = draggedNode;

                // Mark dirty and refresh combo box
                _isDirty = true;
                cmbDefaultFolder.Items.Clear();
                PopulateComboBoxFromTree(tvProjectSubFolders, cmbDefaultFolder);
            }
            finally
            {
                tv.EndUpdate();
            }
        }
    }
}