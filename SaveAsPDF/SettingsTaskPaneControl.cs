using System;
using System.IO;
using System.Windows.Forms;
using System.Drawing;
using SaveAsPDF.Helpers;
using SaveAsPDF.Models;
using SaveAsPDF.Properties;

namespace SaveAsPDF
{
    // Simplified task-pane version of the settings UI, self-contained (no designer).
    public class SettingsTaskPaneControl : UserControl
    {
        private readonly ISettingsRequester _callingForm;
        private readonly SettingsModel _settingsModel = new SettingsModel();
        private TreeNode _mySelectedNode;
        private bool _isDirty;

        // Core controls we actually need in the task pane
        private readonly TreeView tvProjectSubFolders = new TreeView();
        private readonly ComboBox cmbDefaultFolder = new ComboBox();
        private readonly TextBox txtRootFolder = new TextBox();
        private readonly TextBox txtMinAttSize = new TextBox();
        private readonly TextBox txtTreePath = new TextBox();
        private readonly TextBox txtSaveAsPDFFolder = new TextBox();
        private readonly TextBox txtXmlProjectFile = new TextBox();
        private readonly TextBox txtXmlEmployeesFile = new TextBox();
        private readonly TextBox txtProjectRootTag = new TextBox();
        private readonly TextBox txtDateTag = new TextBox();
        private readonly ListBox lsbLastProjects = new ListBox();
        private readonly TextBox txtMaxProjectCount = new TextBox();
        private readonly Button btnOK = new Button();
        private readonly Button btnSaveSettings = new Button();
        private readonly Button bntCancel = new Button();
        private readonly Button btnFolders = new Button();
        private readonly Button btnLoadTreeFile = new Button();
        private readonly Button btnSaveTreeFile = new Button();
        private readonly Button btnLoadDefaultTree = new Button();
        private readonly StatusStrip statusStrip = new StatusStrip();
        private readonly ToolStripStatusLabel toolStripStatusLabel = new ToolStripStatusLabel();
        private readonly ErrorProvider errorProviderSettings = new ErrorProvider();
        private readonly OpenFileDialog openFileDialog = new OpenFileDialog();
        private readonly SaveFileDialog saveFileDialog = new SaveFileDialog();

        public event EventHandler SettingsCommitted;

        public SettingsTaskPaneControl(ISettingsRequester caller)
        {
            _callingForm = caller ?? throw new ArgumentNullException(nameof(caller));

            // Basic layout for task pane - keep it simple
            Dock = DockStyle.Fill;
            AutoScroll = true;

            // Match FormSettings: RTL orientation
            RightToLeft = RightToLeft.Yes;
            // RightToLeftLayout is a Form property, not available on UserControl
            // Keep only RightToLeft which is supported here.

            // Build minimal UI for the task pane instead of copying full FormSettings layout
            BuildLayout();

            // Initialize state from model
            _callingForm.SettingsComplete(_settingsModel);

            InitializeLogic();

            Load += SettingsTaskPaneControl_Load;
        }

        private void BuildLayout()
        {
            // Main vertical layout
            var main = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 3,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                RightToLeft = RightToLeft.Yes
            };
            main.RowStyles.Add(new RowStyle(SizeType.AutoSize));   // Settings group
            main.RowStyles.Add(new RowStyle(SizeType.Percent, 100)); // Project tree group
            main.RowStyles.Add(new RowStyle(SizeType.AutoSize));   // Buttons

            // === Settings section (similar to FormSettings) ===
            var settingsGroup = new GroupBox
            {
                Text = "הגדרות שמירה",
                Dock = DockStyle.Top,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                RightToLeft = RightToLeft.Yes
            };

            var settingsTable = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 2,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                RightToLeft = RightToLeft.Yes
            };
            settingsTable.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 35));
            settingsTable.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 65));

            // Row 0: Root folder + browse button
            var lblRootFolder = new Label { Text = "תיקיית ראשית", AutoSize = true, TextAlign = ContentAlignment.MiddleLeft, Dock = DockStyle.Fill };
            var pnlRoot = new TableLayoutPanel { ColumnCount = 2, Dock = DockStyle.Fill, AutoSize = true, RightToLeft = RightToLeft.Yes };
            pnlRoot.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            pnlRoot.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            txtRootFolder.Dock = DockStyle.Fill;
            txtRootFolder.RightToLeft = RightToLeft.No; // paths LTR
            btnFolders.Text = "...";
            btnFolders.Width = 30;
            pnlRoot.Controls.Add(txtRootFolder, 0, 0);
            pnlRoot.Controls.Add(btnFolders, 1, 0);
            settingsTable.Controls.Add(lblRootFolder, 0, 0);
            settingsTable.Controls.Add(pnlRoot, 1, 0);

            // Row 1: SaveAsPDF folder
            var lblSaveAsPdf = new Label { Text = "תיקיית SaveAsPDF", AutoSize = true, TextAlign = ContentAlignment.MiddleLeft, Dock = DockStyle.Fill };
            txtSaveAsPDFFolder.Dock = DockStyle.Fill;
            txtSaveAsPDFFolder.RightToLeft = RightToLeft.No;
            settingsTable.Controls.Add(lblSaveAsPdf, 0, 1);
            settingsTable.Controls.Add(txtSaveAsPDFFolder, 1, 1);

            // Row 2: Project XML
            var lblXmlProject = new Label { Text = "קובץ XML פרויקט", AutoSize = true, TextAlign = ContentAlignment.MiddleLeft, Dock = DockStyle.Fill };
            txtXmlProjectFile.Dock = DockStyle.Fill;
            txtXmlProjectFile.RightToLeft = RightToLeft.No;
            settingsTable.Controls.Add(lblXmlProject, 0, 2);
            settingsTable.Controls.Add(txtXmlProjectFile, 1, 2);

            // Row 3: Employees XML
            var lblXmlEmployees = new Label { Text = "קובץ XML עובדים", AutoSize = true, TextAlign = ContentAlignment.MiddleLeft, Dock = DockStyle.Fill };
            txtXmlEmployeesFile.Dock = DockStyle.Fill;
            txtXmlEmployeesFile.RightToLeft = RightToLeft.No;
            settingsTable.Controls.Add(lblXmlEmployees, 0, 3);
            settingsTable.Controls.Add(txtXmlEmployeesFile, 1, 3);

            // Row 4: Project root tag
            var lblProjectRoot = new Label { Text = "תגית שורש פרויקט", AutoSize = true, TextAlign = ContentAlignment.MiddleLeft, Dock = DockStyle.Fill };
            txtProjectRootTag.Dock = DockStyle.Fill;
            settingsTable.Controls.Add(lblProjectRoot, 0, 4);
            settingsTable.Controls.Add(txtProjectRootTag, 1, 4);

            // Row 5: Date tag
            var lblDateTag = new Label { Text = "תגית תאריך", AutoSize = true, TextAlign = ContentAlignment.MiddleLeft, Dock = DockStyle.Fill };
            txtDateTag.Dock = DockStyle.Fill;
            settingsTable.Controls.Add(lblDateTag, 0, 5);
            settingsTable.Controls.Add(txtDateTag, 1, 5);

            // Row 6: Min attachment size
            var lblMinAttSize = new Label { Text = "גודל מינימלי לקובץ מצורף", AutoSize = true, TextAlign = ContentAlignment.MiddleLeft, Dock = DockStyle.Fill };
            txtMinAttSize.Dock = DockStyle.Left;
            txtMinAttSize.Width = 80;
            txtMinAttSize.RightToLeft = RightToLeft.No;
            settingsTable.Controls.Add(lblMinAttSize, 0, 6);
            settingsTable.Controls.Add(txtMinAttSize, 1, 6);

            // Row 7: Tree JSON path + load/save/default
            var lblTreePath = new Label { Text = "קובץ תיקיות (JSON)", AutoSize = true, TextAlign = ContentAlignment.MiddleLeft, Dock = DockStyle.Fill };
            var pnlTreeFile = new FlowLayoutPanel { FlowDirection = FlowDirection.RightToLeft, Dock = DockStyle.Fill, AutoSize = true, RightToLeft = RightToLeft.Yes };
            txtTreePath.Width = 180;
            txtTreePath.RightToLeft = RightToLeft.No;
            btnLoadTreeFile.Text = "פתח";
            btnSaveTreeFile.Text = "שמור";
            btnLoadDefaultTree.Text = "ברירת מחדל";
            // buttons right-to-left, textbox on the left
            pnlTreeFile.Controls.Add(btnLoadDefaultTree);
            pnlTreeFile.Controls.Add(btnSaveTreeFile);
            pnlTreeFile.Controls.Add(btnLoadTreeFile);
            pnlTreeFile.Controls.Add(txtTreePath);
            settingsTable.Controls.Add(lblTreePath, 0, 7);
            settingsTable.Controls.Add(pnlTreeFile, 1, 7);

            // Row 8: Default folder combo
            var lblDefaultFolder = new Label { Text = "תיקיית ברירת מחדל בעץ", AutoSize = true, TextAlign = ContentAlignment.MiddleLeft, Dock = DockStyle.Fill };
            cmbDefaultFolder.Dock = DockStyle.Fill;
            settingsTable.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            settingsTable.Controls.Add(lblDefaultFolder, 0, 8);
            settingsTable.Controls.Add(cmbDefaultFolder, 1, 8);

            // Row 9: Last projects list + max count
            var lblLastProjects = new Label { Text = "פרויקטים אחרונים", AutoSize = true, TextAlign = ContentAlignment.MiddleLeft, Dock = DockStyle.Fill };
            lsbLastProjects.Height = 80;
            lsbLastProjects.Dock = DockStyle.Fill;
            settingsTable.RowStyles.Add(new RowStyle(SizeType.Absolute, 80));
            settingsTable.Controls.Add(lblLastProjects, 0, 9);
            settingsTable.Controls.Add(lsbLastProjects, 1, 9);

            var lblMaxProj = new Label { Text = "מס׳ פרויקטים אחרונים", AutoSize = true, TextAlign = ContentAlignment.MiddleLeft, Dock = DockStyle.Fill };
            txtMaxProjectCount.Width = 40;
            txtMaxProjectCount.RightToLeft = RightToLeft.No;
            settingsTable.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            settingsTable.Controls.Add(lblMaxProj, 0, 10);
            settingsTable.Controls.Add(txtMaxProjectCount, 1, 10);

            settingsGroup.Controls.Add(settingsTable);

            // === Project root folder preview (similar to tvFolders in FormMain) ===
            var projectTreeGroup = new GroupBox
            {
                Text = "עץ תיקיות פרויקט",
                Dock = DockStyle.Fill,
                RightToLeft = RightToLeft.Yes
            };
            tvProjectSubFolders.Dock = DockStyle.Fill;
            tvProjectSubFolders.HideSelection = false;
            tvProjectSubFolders.PathSeparator = "\\";
            projectTreeGroup.Controls.Add(tvProjectSubFolders);

            // === Buttons at bottom ===
            var pnlButtons = new FlowLayoutPanel
            {
                FlowDirection = FlowDirection.RightToLeft,
                Dock = DockStyle.Fill,
                AutoSize = true,
                RightToLeft = RightToLeft.Yes
            };
            btnOK.Text = "אישור";
            btnSaveSettings.Text = "שמור";
            bntCancel.Text = "ביטול";
            pnlButtons.Controls.Add(btnOK);
            pnlButtons.Controls.Add(btnSaveSettings);
            pnlButtons.Controls.Add(bntCancel);

            // Status strip
            statusStrip.Items.Add(toolStripStatusLabel);
            statusStrip.Dock = DockStyle.Bottom;
            statusStrip.RightToLeft = RightToLeft.Yes;

            // Assemble main layout
            main.Controls.Add(settingsGroup, 0, 0);
            main.Controls.Add(projectTreeGroup, 0, 1);
            main.Controls.Add(pnlButtons, 0, 2);

            Controls.Add(main);
            Controls.Add(statusStrip);

            errorProviderSettings.ContainerControl = this;
            errorProviderSettings.RightToLeft = true;
        }

        private void InitializeLogic()
        {
            // Drag & drop on tree
            tvProjectSubFolders.AllowDrop = true;
            tvProjectSubFolders.ItemDrag += TvProjectSubFolders_ItemDrag;
            tvProjectSubFolders.DragEnter += TvProjectSubFolders_DragEnter;
            tvProjectSubFolders.DragOver += TvProjectSubFolders_DragOver;
            tvProjectSubFolders.DragDrop += TvProjectSubFolders_DragDrop;

            // Wire button clicks
            btnOK.Click += btnOK_Click;
            btnSaveSettings.Click += btnSaveSettings_Click;
            bntCancel.Click += bntCancel_Click;
            btnFolders.Click += btnFolders_Click;
            btnLoadTreeFile.Click += btnLoadTreeFile_Click;
            btnSaveTreeFile.Click += btnSaveTreeFile_Click;
            btnLoadDefaultTree.Click += btnLoadDefaultTree_Click;

            // Text change / validation events
            txtRootFolder.TextChanged += txtRootFolder_TextChanged;
            txtMinAttSize.TextChanged += txtMinAttSize_TextChanged;
            txtMinAttSize.Validated += txtMinAttSize_Validated;
            txtMinAttSize.Validating += txtMinAttSize_Validating;
            txtTreePath.TextChanged += txtTreePath_TextChanged;
            txtSaveAsPDFFolder.TextChanged += txtSaveAsPDFFolder_TextChanged;
            txtXmlProjectFile.TextChanged += txtXmlProjectFile_TextChanged;
            txtXmlEmployeesFile.TextChanged += txtXmlEmployeesFile_TextChanged;
            txtProjectRootTag.TextChanged += txtProjectRootTag_TextChanged;
            txtDateTag.TextChanged += txtDateTag_TextChanged;
            txtMaxProjectCount.Validated += txtMaxProjectCount_Validated;
            txtMaxProjectCount.Validating += txtMaxProjectCount_Validating;

            cmbDefaultFolder.SelectedIndexChanged += cmbDefaultFolder_SelectedIndexChanged;

            // Status help wiring (minimal)
            WireStatusHelp();
        }

        private void SettingsTaskPaneControl_Load(object sender, EventArgs e)
        {
            txtSaveAsPDFFolder.Text = _settingsModel.XmlSaveAsPDFFolder;
            txtXmlProjectFile.Text = _settingsModel.XmlProjectFile;
            txtXmlEmployeesFile.Text = _settingsModel.XmlEmployeesFile;
            txtProjectRootTag.Text = _settingsModel.ProjectRootTag;
            txtDateTag.Text = _settingsModel.DateTag;

            // Ensure LastProjects is always non-null to avoid NullReferenceException on first use
            if (Settings.Default.LastProjects == null)
            {
                Settings.Default.LastProjects = new System.Collections.Specialized.StringCollection();
                Settings.Default.Save();
            }

            lsbLastProjects.Items.Clear();
            foreach (string item in Settings.Default.LastProjects)
            {
                if (!lsbLastProjects.Items.Contains(item))
                {
                    lsbLastProjects.Items.Add(item);
                }
            }

            txtMaxProjectCount.Text = lsbLastProjects.Items.Count.ToString();
            _isDirty = false;
        }

        private void ControlMouseEnterStatus(object sender, EventArgs e)
        {
            var ctl = sender as Control;
            toolStripStatusLabel.Text = ctl?.Tag as string ?? string.Empty;
        }

        private void ControlMouseLeaveStatus(object sender, EventArgs e)
        {
            toolStripStatusLabel.Text = string.Empty;
        }

        private void WireStatusHelp()
        {
            btnOK.Tag = "אישור";
            btnOK.MouseEnter += ControlMouseEnterStatus;
            btnOK.MouseLeave += ControlMouseLeaveStatus;

            btnSaveSettings.Tag = "שמור שינויים";
            btnSaveSettings.MouseEnter += ControlMouseEnterStatus;
            btnSaveSettings.MouseLeave += ControlMouseLeaveStatus;

            bntCancel.Tag = "בטל ושחזר";
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
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            if (_isDirty)
            {
                DialogResult result = XMessageBox.Show(
                    "שמור שינויים?",
                    "SaveAsPDF",
                    XMessageBoxButtons.YesNo,
                    XMessageBoxIcon.Question,
                    XMessageAlignment.Right,
                    XMessageLanguage.Hebrew
                );
                if (result == DialogResult.Yes)
                {
                    if (cmbDefaultFolder.SelectedItem != null)
                    {
                        string selectedPath = cmbDefaultFolder.SelectedItem.ToString();
                        _settingsModel.DefaultSavePath = selectedPath;
                        _settingsModel.DefaultFolderID = cmbDefaultFolder.SelectedIndex;
                    }
                    SettingsHelpers.SaveModelToSettings(_settingsModel);
                }
            }
            _callingForm.SettingsComplete(_settingsModel);
            _isDirty = false;
            SettingsCommitted?.Invoke(this, EventArgs.Empty);
        }

        private void btnSaveSettings_Click(object sender, EventArgs e)
        {
            if (cmbDefaultFolder.SelectedItem != null)
            {
                string selectedPath = cmbDefaultFolder.SelectedItem.ToString();
                _settingsModel.DefaultSavePath = selectedPath;
                _settingsModel.DefaultFolderID = cmbDefaultFolder.SelectedIndex;
            }

            SettingsHelpers.SaveModelToSettings(_settingsModel);
            _isDirty = false;
        }

        private void bntCancel_Click(object sender, EventArgs e)
        {
            ReloadFromSettings();
            _isDirty = false;
        }

        private void ReloadFromSettings()
        {
            var reloaded = SettingsHelpers.LoadProjectSettings();
            if (reloaded != null)
            {
                _settingsModel.RootDrive = reloaded.RootDrive;
                _settingsModel.DefaultSavePath = reloaded.DefaultSavePath;
                _settingsModel.DefaultFolderID = reloaded.DefaultFolderID;
                _settingsModel.XmlSaveAsPDFFolder = reloaded.XmlSaveAsPDFFolder;
                _settingsModel.XmlProjectFile = reloaded.XmlProjectFile;
                _settingsModel.XmlEmployeesFile = reloaded.XmlEmployeesFile;
                _settingsModel.ProjectRootTag = reloaded.ProjectRootTag;
                _settingsModel.DateTag = reloaded.DateTag;
                _settingsModel.DefaultTreeFile = reloaded.DefaultTreeFile;
                _settingsModel.MinAttachmentSize = reloaded.MinAttachmentSize;

                txtRootFolder.Text = _settingsModel.RootDrive;
                txtSaveAsPDFFolder.Text = _settingsModel.XmlSaveAsPDFFolder;
                txtXmlProjectFile.Text = _settingsModel.XmlProjectFile;
                txtXmlEmployeesFile.Text = _settingsModel.XmlEmployeesFile;
                txtProjectRootTag.Text = _settingsModel.ProjectRootTag;
                txtDateTag.Text = _settingsModel.DateTag;
                txtTreePath.Text = _settingsModel.DefaultTreeFile;
                txtMinAttSize.Text = _settingsModel.MinAttachmentSize.ToString();
            }
        }

        private void PopulateComboBoxFromTree(TreeView treeView, ComboBox comboBox)
        {
            comboBox.Items.Clear();
            if (treeView == null) return;

            if (string.IsNullOrEmpty(treeView.PathSeparator))
                treeView.PathSeparator = "\\";

            var stack = new System.Collections.Generic.Stack<TreeNode>();
            for (int r = treeView.Nodes.Count - 1; r >= 0; r--)
                stack.Push(treeView.Nodes[r]);

            while (stack.Count > 0)
            {
                var node = stack.Pop();
                if (node == null) continue;

                string fullPath = node.FullPath ?? string.Empty;
                if (!string.IsNullOrEmpty(fullPath))
                    comboBox.Items.Add(fullPath);

                if (node.Nodes != null && node.Nodes.Count > 0)
                {
                    for (int i = node.Nodes.Count - 1; i >= 0; i--)
                        stack.Push(node.Nodes[i]);
                }
            }
        }

        private void TvProjectSubFolders_ItemDrag(object sender, ItemDragEventArgs e)
        {
            if (e.Item is TreeNode node)
                DoDragDrop(node, DragDropEffects.Move);
        }

        private void TvProjectSubFolders_DragEnter(object sender, DragEventArgs e)
        {
            e.Effect = e.Data.GetDataPresent(typeof(TreeNode)) ? DragDropEffects.Move : DragDropEffects.None;
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

            if (targetNode != null)
            {
                TreeNode cur = targetNode;
                while (cur != null)
                {
                    if (cur == draggedNode) return;
                    cur = cur.Parent;
                }
            }

            tv.BeginUpdate();
            try
            {
                if (draggedNode.Parent != null)
                    draggedNode.Parent.Nodes.Remove(draggedNode);
                else
                    tv.Nodes.Remove(draggedNode);

                if (targetNode == null)
                    tv.Nodes.Add(draggedNode);
                else
                    targetNode.Nodes.Add(draggedNode);

                draggedNode.EnsureVisible();
                tv.SelectedNode = draggedNode;

                _isDirty = true;
                cmbDefaultFolder.Items.Clear();
                PopulateComboBoxFromTree(tvProjectSubFolders, cmbDefaultFolder);
            }
            finally
            {
                tv.EndUpdate();
            }
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
            openFileDialog.Title = "פתח קובץ מבנה (JSON)";
            openFileDialog.Filter = "JSON tree (*.json)|*.json";

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                var file = openFileDialog.FileName;
                try
                {
                    JsonHelper.LoadTreeFromJson(file, tvProjectSubFolders);
                    txtTreePath.Text = file;
                    _settingsModel.DefaultTreeFile = file;
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

                    JsonHelper.SaveTreeToJson(tvProjectSubFolders, file);

                    tvProjectSubFolders.ExpandAll();
                    txtTreePath.Text = file;
                    _settingsModel.DefaultTreeFile = file;
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

        private void btnLoadDefaultTree_Click(object sender, EventArgs e)
        {
            tvProjectSubFolders.Nodes.Clear();
            if (!string.IsNullOrEmpty(_settingsModel.ProjectRootTag))
            {
                tvProjectSubFolders.Nodes.Add(_settingsModel.ProjectRootTag);
            }
            cmbDefaultFolder.Items.Clear();
            PopulateComboBoxFromTree(tvProjectSubFolders, cmbDefaultFolder);
            _isDirty = true;
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

        private void cmbDefaultFolder_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cmbDefaultFolder.SelectedItem != null)
            {
                string selectedPath = cmbDefaultFolder.SelectedItem.ToString();
                _settingsModel.DefaultSavePath = selectedPath;
                _settingsModel.DefaultFolderID = cmbDefaultFolder.SelectedIndex;
                _isDirty = true;

                SettingsHelpers.SaveModelToSettings(_settingsModel);
                _callingForm.SettingsComplete(_settingsModel);
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

        private void txtMinAttSize_TextChanged(object sender, EventArgs e)
        {
            _isDirty = true;
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
