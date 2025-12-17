// Task-pane version of FormMain with full right-to-left layout and core logic
using Microsoft.Office.Interop.Outlook;
using SaveAsPDF.Helpers;
using SaveAsPDF.Models;
using SaveAsPDF.Properties;
using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using System.Drawing;
using Exception = System.Exception;

namespace SaveAsPDF
{
    public class MainFormTaskPaneControl : UserControl, IEmployeeRequester, INewProjectRequester, ISettingsRequester
    {
        private ProjectModel _projectModel = new ProjectModel();
        public static SettingsModel settingsModel = new SettingsModel();

        private MailItem _mailItem;
        private string _currentProjectId;
        private bool _isDoubleClick;

        private readonly List<string> searchHistory = new List<string>();
        private readonly List<AttachmentsModel> attachmentsModels = new List<AttachmentsModel>();
        private readonly BindingList<EmployeeModel> _employeesBindingList = new BindingList<EmployeeModel>();
        private bool _selectingLeader;

        // Core controls roughly mirroring FormMain, but as a task pane
        private readonly TextBox txtProjectID = new TextBox();
        private readonly TextBox txtProjectName = new TextBox();
        private readonly TextBox txtProjectLeader = new TextBox();
        private readonly TextBox txtSubject = new TextBox();
        private readonly TextBox txtFullPath = new TextBox();
        private readonly ComboBox cmbSaveLocation = new ComboBox();
        private readonly RichTextBox rtxtNotes = new RichTextBox();
        private readonly RichTextBox rtxtProjectNotes = new RichTextBox();
        private readonly DataGridView dgvAttachments = new DataGridView();
        private readonly DataGridView dgvEmployees = new DataGridView();
        private readonly TreeView tvFolders = new TreeView();
        private readonly CheckBox chkbSendNote = new CheckBox();
        private readonly CheckBox chkbSelectAllAttachments = new CheckBox();
        private readonly CheckBox chbOpenPDF = new CheckBox();
        private readonly Button btnOK = new Button();
        private readonly Button btnCancel = new Button();
        private readonly Button btnSettings = new Button();
        private readonly Button btnNewProject = new Button();
        private readonly Button btnFolders = new Button();
        private readonly Button btnCopyNotesToMail = new Button();
        private readonly Button btnCopyNotesToProject = new Button();
        private readonly Button btnStyle = new Button();
        private readonly Button btnRemoveEmployee = new Button();
        private readonly Button btnPhoneBook = new Button();
        private readonly Button btnProjectLeader = new Button();
        private readonly TabControl tabNotes = new TabControl();
        private readonly TabControl tabFilesFolders = new TabControl();
        private readonly StatusStrip statusStrip = new StatusStrip();
        private readonly ToolStripStatusLabel tsslStatus = new ToolStripStatusLabel();
        private readonly ErrorProvider errorProviderMain = new ErrorProvider();
        private readonly FontDialog dlgFont = new FontDialog();

        public MainFormTaskPaneControl()
        {
            Dock = DockStyle.Fill;
            AutoScroll = true;

            RightToLeft = RightToLeft.Yes;

            BuildLayout();
            ConfigureEmployeeDataGrid();

            KeyDown += MainFormTaskPaneControl_KeyDown;

            btnProjectLeader.Click += btnProjectLeader_Click;
            btnSettings.CausesValidation = false;
            btnCancel.CausesValidation = false;

            Load += MainFormTaskPaneControl_Load;
        }

        // Called by ThisAddIn when selection changes
        public void LoadMailItem(MailItem mailItem)
        {
            // Ignore if the selection did not actually change
            if (ReferenceEquals(_mailItem, mailItem))
            {
                return;
            }

            // Clear UI when there is no selected mail
            if (mailItem == null)
            {
                _mailItem = null;
                ClearMailRelatedUi();
                return;
            }

            _mailItem = mailItem;

            // Derive a project id from the mail subject if possible
            string subject = _mailItem.Subject ?? string.Empty;
            string projectIdFromSubject = ExtractProjectIdFromSubject(subject);

            // Only reload project data if project id actually changed
            if (!string.Equals(_currentProjectId, projectIdFromSubject, StringComparison.Ordinal))
            {
                _currentProjectId = projectIdFromSubject;
                txtProjectID.Text = _currentProjectId;

                if (!string.IsNullOrWhiteSpace(_currentProjectId) && _currentProjectId.SafeProjectID())
                {
                    // Trigger the same flow as manual entry
                    ValidateAndLoadProjectById(_currentProjectId);
                }
                else
                {
                    // If we cannot parse project id, just clear project‑specific UI
                    ClearProjectRelatedUi();
                }
            }

            // Always update subject and attachments for the new mail item
            txtSubject.Text = LoadEmailSubject();
            ProcessMailItem(_mailItem);
        }

        // Simple extraction: look for the first token that looks like a valid project id
        private string ExtractProjectIdFromSubject(string subject)
        {
            if (string.IsNullOrWhiteSpace(subject))
                return string.Empty;

            // Search both subject and body for a token that looks like a valid project id.
            // We rely only on SafeProjectID() to decide, and if nothing matches we return empty
            // so the caller can leave txtProjectID blank without showing an error.
            string combinedText = subject;
            try
            {
                if (_mailItem != null)
                {
                    string body = _mailItem.Body ?? string.Empty;
                    if (!string.IsNullOrWhiteSpace(body))
                    {
                        combinedText = subject + " " + body;
                    }
                }
            }
            catch
            {
                // If reading Body fails for any reason, fall back to subject-only search
            }

            var separators = new[] { ' ', '\t', '-', '_', ':', ';', ',', '[', ']', '(', ')', '{', '}' };
            var parts = combinedText.Split(separators, StringSplitOptions.RemoveEmptyEntries);
            foreach (var part in parts)
            {
                if (part.SafeProjectID())
                    return part;
            }

            // No valid project id found in subject or body
            return string.Empty;
        }

        private void ClearMailRelatedUi()
        {
            txtSubject.Clear();
            dgvAttachments.DataSource = null;
            attachmentsModels.Clear();
        }

        private void ClearProjectRelatedUi()
        {
            txtProjectName.Clear();
            txtFullPath.Clear();
            cmbSaveLocation.Items.Clear();
            cmbSaveLocation.Text = string.Empty;
            rtxtProjectNotes.Clear();
            tvFolders.Nodes.Clear();
        }

        private void ValidateAndLoadProjectById(string projectID)
        {
            // Validate project id pattern
            if (!projectID.SafeProjectID())
            {
                errorProviderMain.SetError(txtProjectID, "מספר פרויקט לא חוקי");
                tsslStatus.Text = errorProviderMain.GetError(txtProjectID);
                return;
            }

            errorProviderMain.SetError(txtProjectID, string.Empty);
            tsslStatus.Text = string.Empty;

            // Update history/autocomplete
            UpdateAutoCompleteSource(projectID);

            // Check project folder existence
            var projectRootFolder = projectID.ProjectFullPath(settingsModel.RootDrive);
            if (!Directory.Exists(projectRootFolder.FullName))
            {
                XMessageBox.Show(
                    "הפרויקט לא קיים",
                    "שגיאה",
                    XMessageBoxButtons.OK,
                    XMessageBoxIcon.Error,
                    XMessageAlignment.Right,
                    XMessageLanguage.Hebrew);
                ClearProjectRelatedUi();
                return;
            }

            // Load settings and basic project model for this id
            settingsModel = SettingsHelpers.LoadProjectSettings(projectID);

            // Basic UI update from settings/model
            txtProjectName.Text = _projectModel.ProjectName;
            rtxtProjectNotes.Text = _projectModel.ProjectNotes;

            // Update folder tree root to project root if available
            if (settingsModel.ProjectRootFolder != null && settingsModel.ProjectRootFolder.Exists)
            {
                tvFolders.Nodes.Clear();
                var root = new TreeNode(settingsModel.ProjectRootFolder.Name)
                {
                    Tag = settingsModel.ProjectRootFolder.FullName
                };
                root.Nodes.Add("...");
                tvFolders.Nodes.Add(root);
                txtFullPath.Text = settingsModel.ProjectRootFolder.FullName;
            }
        }

        private void BuildLayout()
        {
            var main = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 3,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                RightToLeft = RightToLeft.Yes
            };
            main.RowStyles.Add(new RowStyle(SizeType.AutoSize)); // top
            main.RowStyles.Add(new RowStyle(SizeType.Percent, 100)); // middle
            main.RowStyles.Add(new RowStyle(SizeType.AutoSize)); // bottom

            // --- Top group: project + subject ---
            var topGroup = new GroupBox
            {
                Text = "נתוני פרויקט והודעה",
                Dock = DockStyle.Top,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                RightToLeft = RightToLeft.Yes
            };

            var topTable = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 2,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                RightToLeft = RightToLeft.Yes
            };
            topTable.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 30));
            topTable.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 70));

            // Project ID
            var lblProjectID = new Label
            {
                Text = "מספר פרויקט",
                AutoSize = true,
                TextAlign = ContentAlignment.MiddleLeft,
                Dock = DockStyle.Fill
            };
            txtProjectID.Dock = DockStyle.Fill;
            txtProjectID.KeyDown += txtProjectID_KeyDown;
            txtProjectID.Validating += txtProjectID_Validating;
            txtProjectID.Validated += txtProjectID_Validated;
            topTable.Controls.Add(lblProjectID, 0, 0);
            topTable.Controls.Add(txtProjectID, 1, 0);

            // Project name
            var lblProjectName = new Label
            {
                Text = "שם הפרויקט",
                AutoSize = true,
                TextAlign = ContentAlignment.MiddleLeft,
                Dock = DockStyle.Fill
            };
            txtProjectName.Dock = DockStyle.Fill;
            topTable.Controls.Add(lblProjectName, 0, 1);
            topTable.Controls.Add(txtProjectName, 1, 1);

            // Project leader
            var lblLeader = new Label
            {
                Text = "מתכנן מוביל",
                AutoSize = true,
                TextAlign = ContentAlignment.MiddleLeft,
                Dock = DockStyle.Fill
            };
            var pnlLeader = new FlowLayoutPanel
            {
                FlowDirection = FlowDirection.LeftToRight,
                Dock = DockStyle.Fill,
                AutoSize = true
            };
            txtProjectLeader.Width = 160;
            btnProjectLeader.Text = "...";
            btnProjectLeader.Width = 30;
            pnlLeader.Controls.Add(txtProjectLeader);
            pnlLeader.Controls.Add(btnProjectLeader);
            topTable.Controls.Add(lblLeader, 0, 2);
            topTable.Controls.Add(pnlLeader, 1, 2);

            // Subject
            var lblSubject = new Label
            {
                Text = "נושא ההודעה",
                AutoSize = true,
                TextAlign = ContentAlignment.MiddleLeft,
                Dock = DockStyle.Fill
            };
            txtSubject.Dock = DockStyle.Fill;
            topTable.Controls.Add(lblSubject, 0, 3);
            topTable.Controls.Add(txtSubject, 1, 3);

            // Full path
            var lblFullPath = new Label
            {
                Text = "נתיב מלא",
                AutoSize = true,
                TextAlign = ContentAlignment.MiddleLeft,
                Dock = DockStyle.Fill
            };
            txtFullPath.Dock = DockStyle.Fill;
            txtFullPath.ReadOnly = true;
            topTable.Controls.Add(lblFullPath, 0, 4);
            topTable.Controls.Add(txtFullPath, 1, 4);

            // Save location combo + folder selector
            var lblSaveLocation = new Label
            {
                Text = "מיקום שמירה",
                AutoSize = true,
                TextAlign = ContentAlignment.MiddleLeft,
                Dock = DockStyle.Fill
            };
            var pnlSaveLocation = new FlowLayoutPanel
            {
                FlowDirection = FlowDirection.LeftToRight,
                Dock = DockStyle.Fill,
                AutoSize = true
            };
            cmbSaveLocation.Width = 220;
            cmbSaveLocation.SelectedValueChanged += cmbSaveLocation_SelectedValueChanged;
            cmbSaveLocation.TextUpdate += cmbSaveLocation_TextUpdate;
            btnFolders.Text = "בחר תיקייה";
            btnFolders.Click += btnFolders_Click;
            pnlSaveLocation.Controls.Add(cmbSaveLocation);
            pnlSaveLocation.Controls.Add(btnFolders);
            topTable.Controls.Add(lblSaveLocation, 0, 5);
            topTable.Controls.Add(pnlSaveLocation, 1, 5);

            topGroup.Controls.Add(topTable);

            // --- Middle: notes, attachments, employees, folders ---
            var middleSplit = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 2,
                RowCount = 1,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                RightToLeft = RightToLeft.Yes
            };
            middleSplit.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
            middleSplit.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));

            // Left: notes tabs
            tabNotes.Dock = DockStyle.Fill;
            var tabMailNotes = new TabPage("הערות למייל") { RightToLeft = RightToLeft.Yes };
            var tabProjectNotes = new TabPage("הערות בפרויקט") { RightToLeft = RightToLeft.Yes };
            rtxtNotes.Dock = DockStyle.Fill;
            rtxtProjectNotes.Dock = DockStyle.Fill;
            tabMailNotes.Controls.Add(rtxtNotes);
            tabProjectNotes.Controls.Add(rtxtProjectNotes);
            tabNotes.TabPages.Add(tabMailNotes);
            tabNotes.TabPages.Add(tabProjectNotes);

            var pnlNotesWithButtons = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                RowCount = 2,
                ColumnCount = 1,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink
            };
            pnlNotesWithButtons.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            pnlNotesWithButtons.RowStyles.Add(new RowStyle(SizeType.AutoSize));

            var pnlNotesButtons = new FlowLayoutPanel
            {
                FlowDirection = FlowDirection.RightToLeft,
                Dock = DockStyle.Fill,
                AutoSize = true
            };
            btnCopyNotesToMail.Text = "העתק לפרויקט ← מייל";
            btnCopyNotesToMail.Click += btnCopyNotesToMail_Click;
            btnCopyNotesToProject.Text = "העתק ממייל → פרויקט";
            btnCopyNotesToProject.Click += btnCopyNotesToProject_Click;
            btnStyle.Text = "גופן";
            btnStyle.Click += btnStyle_Click;
            pnlNotesButtons.Controls.Add(btnCopyNotesToMail);
            pnlNotesButtons.Controls.Add(btnCopyNotesToProject);
            pnlNotesButtons.Controls.Add(btnStyle);

            pnlNotesWithButtons.Controls.Add(tabNotes, 0, 0);
            pnlNotesWithButtons.Controls.Add(pnlNotesButtons, 0, 1);

            // Right: attachments + employees + folders
            var rightSplit = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 3,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink
            };
            rightSplit.RowStyles.Add(new RowStyle(SizeType.Percent, 40));
            rightSplit.RowStyles.Add(new RowStyle(SizeType.Percent, 30));
            rightSplit.RowStyles.Add(new RowStyle(SizeType.Percent, 30));

            // Attachments
            var attachmentsGroup = new GroupBox
            {
                Text = "קבצים מצורפים",
                Dock = DockStyle.Fill,
                RightToLeft = RightToLeft.Yes
            };
            var attachmentsTable = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                RowCount = 2,
                ColumnCount = 1
            };
            attachmentsTable.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            attachmentsTable.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            chkbSelectAllAttachments.Text = "הסר הכל";
            chkbSelectAllAttachments.Checked = true;
            chkbSelectAllAttachments.CheckedChanged += chkbSelectAllAttachments_CheckedChanged;
            attachmentsTable.Controls.Add(chkbSelectAllAttachments, 0, 0);

            dgvAttachments.Dock = DockStyle.Fill;
            dgvAttachments.CellDoubleClick += dgvAttachments_CellDoubleClick;
            attachmentsTable.Controls.Add(dgvAttachments, 0, 1);
            attachmentsGroup.Controls.Add(attachmentsTable);

            // Employees
            var employeesGroup = new GroupBox
            {
                Text = "עובדים בפרויקט",
                Dock = DockStyle.Fill,
                RightToLeft = RightToLeft.Yes
            };
            var employeesTable = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                RowCount = 2,
                ColumnCount = 2
            };
            employeesTable.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            employeesTable.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            employeesTable.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
            employeesTable.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));

            btnPhoneBook.Text = "ספר טלפונים";
            btnPhoneBook.Click += btnPhoneBook_Click;
            btnRemoveEmployee.Text = "הסר עובד";
            btnRemoveEmployee.Click += btnRemoveEmployee_Click;

            var pnlEmpButtons = new FlowLayoutPanel
            {
                FlowDirection = FlowDirection.RightToLeft,
                Dock = DockStyle.Fill,
                AutoSize = true
            };
            pnlEmpButtons.Controls.Add(btnPhoneBook);
            pnlEmpButtons.Controls.Add(btnRemoveEmployee);

            employeesTable.Controls.Add(pnlEmpButtons, 0, 0);
            employeesTable.SetColumnSpan(pnlEmpButtons, 2);

            dgvEmployees.Dock = DockStyle.Fill;
            dgvEmployees.CurrentCellDirtyStateChanged += dgvEmployees_CurrentCellDirtyStateChanged;
            dgvEmployees.CellValueChanged += dgvEmployees_CellValueChanged;
            employeesTable.Controls.Add(dgvEmployees, 0, 1);
            employeesTable.SetColumnSpan(dgvEmployees, 2);

            employeesGroup.Controls.Add(employeesTable);

            // Folders
            var foldersGroup = new GroupBox
            {
                Text = "עץ תיקיות פרויקט",
                Dock = DockStyle.Fill,
                RightToLeft = RightToLeft.Yes
            };
            tvFolders.Dock = DockStyle.Fill;
            tvFolders.PathSeparator = "\\";
            tvFolders.BeforeExpand += tvFolders_BeforeExpand;
            tvFolders.BeforeCollapse += tvFolders_BeforeCollapse;
            tvFolders.MouseDown += tvFolders_MouseDown;
            tvFolders.NodeMouseClick += tvFolders_NodeMouseClick;
            tvFolders.NodeMouseDoubleClick += tvFolders_NodeMouseDoubleClick;
            foldersGroup.Controls.Add(tvFolders);

            rightSplit.Controls.Add(attachmentsGroup, 0, 0);
            rightSplit.Controls.Add(employeesGroup, 0, 1);
            rightSplit.Controls.Add(foldersGroup, 0, 2);

            middleSplit.Controls.Add(pnlNotesWithButtons, 0, 0);
            middleSplit.Controls.Add(rightSplit, 1, 0);

            // --- Bottom buttons + status ---
            var bottomPanel = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 2,
                AutoSize = true
            };
            bottomPanel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            bottomPanel.RowStyles.Add(new RowStyle(SizeType.AutoSize));

            var pnlButtons = new FlowLayoutPanel
            {
                FlowDirection = FlowDirection.RightToLeft,
                Dock = DockStyle.Fill,
                AutoSize = true
            };
            btnOK.Text = "שמור ל-PDF";
            btnOK.Click += btnOK_Click;
            btnCancel.Text = "בטל";
            btnCancel.Click += btnCancel_Click;
            btnSettings.Text = "הגדרות";
            btnSettings.Click += BtnSettings_Click;
            btnNewProject.Text = "פרויקט חדש";
            btnNewProject.Click += btnNewProject_Click;
            chbOpenPDF.Text = "פתח PDF לאחר שמירה";
            chbOpenPDF.CheckedChanged += chbOpenPDF_CheckedChanged;

            pnlButtons.Controls.Add(btnOK);
            pnlButtons.Controls.Add(btnCancel);
            pnlButtons.Controls.Add(btnSettings);
            pnlButtons.Controls.Add(btnNewProject);
            pnlButtons.Controls.Add(chbOpenPDF);

            statusStrip.Items.Add(tsslStatus);
            statusStrip.RightToLeft = RightToLeft.Yes;

            bottomPanel.Controls.Add(pnlButtons, 0, 0);
            bottomPanel.Controls.Add(statusStrip, 0, 1);

            main.Controls.Add(topGroup, 0, 0);
            main.Controls.Add(middleSplit, 0, 1);
            main.Controls.Add(bottomPanel, 0, 2);

            Controls.Add(main);

            errorProviderMain.ContainerControl = this;
            errorProviderMain.RightToLeft = true;

            WireStatusHelp();
        }

        private void MainFormTaskPaneControl_Load(object sender, EventArgs e)
        {
            SetupContextMenus();
            chkbSelectAllAttachments.Checked = true;
            chkbSelectAllAttachments.Text = "הסר הכל";
            ConfigureProjectIdAutoComplete();
            LoadSearchHistory();
            settingsModel = SettingsHelpers.LoadProjectSettings();
            txtSubject.Text = LoadEmailSubject();
            PopulateDriveNodes();
        }

        private string LoadEmailSubject() => _mailItem?.Subject ?? "Default Email Subject";

        private void SetupContextMenus()
        {
            rtxtNotes.EnableContextMenu();
            rtxtProjectNotes.EnableContextMenu();
            txtFullPath.EnableContextMenu();
            txtProjectID.EnableContextMenu();
            txtProjectName.EnableContextMenu();
            tvFolders.EnableContextMenu();
        }

        private void ConfigureProjectIdAutoComplete()
        {
            txtProjectID.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            txtProjectID.AutoCompleteSource = AutoCompleteSource.CustomSource;
        }

        private void PopulateDriveNodes()
        {
            tvFolders.Nodes.Clear();

            foreach (string drive in Environment.GetLogicalDrives())
            {
                try
                {
                    var di = new DriveInfo(drive);
                    var node = new TreeNode(drive.Substring(0, 1)) { Tag = drive };
                    if (di.IsReady)
                        node.Nodes.Add("...");
                    tvFolders.Nodes.Add(node);
                }
                catch
                {
                }
            }
        }

        private void UpdateAutoCompleteSource(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
                return;

            searchHistory.Remove(text);
            searchHistory.Insert(0, text);

            int maxCount = 10;
            if (searchHistory.Count > maxCount)
                searchHistory.RemoveRange(maxCount, searchHistory.Count - maxCount);

            if (Settings.Default.LastProjects == null)
                Settings.Default.LastProjects = new StringCollection();
            Settings.Default.LastProjects.Clear();
            Settings.Default.LastProjects.AddRange(searchHistory.ToArray());
            Settings.Default.Save();

            txtProjectID.AutoCompleteCustomSource.Clear();
            txtProjectID.AutoCompleteCustomSource.AddRange(searchHistory.ToArray());
        }

        private void LoadSearchHistory()
        {
            if (Settings.Default.LastProjects == null)
                Settings.Default.LastProjects = new StringCollection();
            searchHistory.Clear();
            searchHistory.AddRange(Settings.Default.LastProjects.Cast<string>().Distinct());
            txtProjectID.AutoCompleteCustomSource.Clear();
            txtProjectID.AutoCompleteCustomSource.AddRange(searchHistory.ToArray());
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
        }

        private void btnFolders_Click(object sender, EventArgs e)
        {
            var dialog = new FolderPicker
            {
                InputPath = settingsModel.ProjectRootFolder != null && settingsModel.ProjectRootFolder.Exists
                    ? settingsModel.ProjectRootFolder.FullName
                    : settingsModel.RootDrive
            };
            if (dialog.ShowDialog(Handle) == true)
            {
                tvFolders.Nodes.Clear();
                var root = new TreeNode(new DirectoryInfo(dialog.ResultPath).Name) { Tag = dialog.ResultPath };
                root.Nodes.Add("...");
                tvFolders.Nodes.Add(root);
            }
        }

        private void txtProjectID_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                btnOK.Select();
        }

        private void txtProjectID_Validating(object sender, System.ComponentModel.CancelEventArgs e)
        {
            if (!txtProjectID.Text.SafeProjectID())
            {
                errorProviderMain.SetError(txtProjectID, "מספר פרויקט לא חוקי");
                txtProjectID.Select(0, txtProjectID.Text.Length);
                tsslStatus.Text = errorProviderMain.GetError(txtProjectID);
                e.Cancel = true;
            }
        }

        private void txtProjectID_Validated(object sender, EventArgs e)
        {
            errorProviderMain.SetError(txtProjectID, string.Empty);
            tsslStatus.Text = errorProviderMain.GetError(txtProjectID);

            string projectID = txtProjectID.Text;
            if (!string.IsNullOrWhiteSpace(projectID))
            {
                ValidateAndLoadProjectById(projectID);
            }
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            XMessageBox.Show("MainFormTaskPaneControl OK clicked (stub)", "SaveAsPDF");
        }

        private void BtnSettings_Click(object sender, EventArgs e)
        {
            Globals.ThisAddIn.ToggleSettingsPane();
        }

        private void btnNewProject_Click(object sender, EventArgs e)
        {
            using (var frm = new FormNewProject(this))
            {
                frm.ShowDialog();
            }
        }

        private void chbOpenPDF_CheckedChanged(object sender, EventArgs e)
        {
            settingsModel.OpenPDF = chbOpenPDF.Checked;
        }

        private void MainFormTaskPaneControl_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
            {
                txtProjectID.Clear();
            }
        }

        private void ConfigureEmployeeDataGrid()
        {
            dgvEmployees.AutoGenerateColumns = false;
            dgvEmployees.Columns.Clear();

            dgvEmployees.Columns.AddRange(new DataGridViewColumn[]
            {
                new DataGridViewTextBoxColumn
                {
                    Name = "Id",
                    DataPropertyName = "Id",
                    Visible = false
                },
                new DataGridViewTextBoxColumn
                {
                    Name = "FirstName",
                    DataPropertyName = "FirstName",
                    HeaderText = "שם פרטי"
                },
                new DataGridViewTextBoxColumn
                {
                    Name = "LastName",
                    DataPropertyName = "LastName",
                    HeaderText = "שם משפחה"
                },
                new DataGridViewTextBoxColumn
                {
                    Name = "EmailAddress",
                    DataPropertyName = "EmailAddress",
                    HeaderText = "אימייל"
                }
            });

            dgvEmployees.DataSource = _employeesBindingList;
            dgvEmployees.ReadOnly = true;
        }

        public void SettingsComplete(SettingsModel settings)
        {
            settingsModel = SettingsHelpers.LoadSettingsToModel(settings);
        }

        public void EmployeeComplete(EmployeeModel model)
        {
            if (!_employeesBindingList.Any(e => e.EmailAddress == model.EmailAddress))
            {
                _employeesBindingList.Add(model);
            }
        }

        public void NewProjectComplete(ProjectModel model)
        {
            _projectModel = model;
            txtProjectID.Text = _projectModel.ProjectNumber;
            txtProjectName.Text = _projectModel.ProjectName;
            rtxtProjectNotes.Text = _projectModel.ProjectNotes;
        }

        private void btnProjectLeader_Click(object sender, EventArgs e)
        {
            _selectingLeader = true;
            using (var frmContacts = new FormContacts(this))
            {
                frmContacts.ShowDialog();
            }
            _selectingLeader = false;
        }

        private void ProcessMailItem(MailItem mailItem)
        {
            txtSubject.Text = mailItem.Subject;
        }

        // Status strip hover helpers adapted from FormMain
        private void MouseEnterStatus(object sender, EventArgs e)
        {
            var ctl = sender as Control;
            if (ctl != null)
                tsslStatus.Text = ctl.Tag as string ?? string.Empty;
        }

        private void MouseLeaveStatus(object sender, EventArgs e)
        {
            tsslStatus.Text = string.Empty;
        }

        private void WireStatusHelp()
        {
            btnOK.Tag = "שמור ל-PDF";
            btnOK.MouseEnter += MouseEnterStatus; btnOK.MouseLeave += MouseLeaveStatus;

            btnCancel.Tag = "בטל וסגור";
            btnCancel.MouseEnter += MouseEnterStatus; btnCancel.MouseLeave += MouseLeaveStatus;

            btnSettings.Tag = "הגדרות";
            btnSettings.MouseEnter += MouseEnterStatus; btnSettings.MouseLeave += MouseLeaveStatus;

            btnNewProject.Tag = "פרויקט חדש";
            btnNewProject.MouseEnter += MouseEnterStatus; btnNewProject.MouseLeave += MouseLeaveStatus;

            btnFolders.Tag = "בחר תיקיית שורש";
            btnFolders.MouseEnter += MouseEnterStatus; btnFolders.MouseLeave += MouseLeaveStatus;

            btnCopyNotesToMail.Tag = "העתק הערות לפרויקט אל המייל";
            btnCopyNotesToMail.MouseEnter += MouseEnterStatus; btnCopyNotesToMail.MouseLeave += MouseLeaveStatus;

            btnCopyNotesToProject.Tag = "העתק ערות מהמייל אל הפרויקט";
            btnCopyNotesToProject.MouseEnter += MouseEnterStatus; btnCopyNotesToProject.MouseLeave += MouseLeaveStatus;

            btnStyle.Tag = "בחר גופן להערות";
            btnStyle.MouseEnter += MouseEnterStatus; btnStyle.MouseLeave += MouseLeaveStatus;

            btnRemoveEmployee.Tag = "הסר עובד מהרשימה";
            btnRemoveEmployee.MouseEnter += MouseEnterStatus; btnRemoveEmployee.MouseLeave += MouseLeaveStatus;

            btnPhoneBook.Tag = "בחר עובד מספר טלפונים";
            btnPhoneBook.MouseEnter += MouseEnterStatus; btnPhoneBook.MouseLeave += MouseLeaveStatus;

            btnProjectLeader.Tag = "בחר מתכנן מוביל בפרויקט";
            btnProjectLeader.MouseEnter += MouseEnterStatus; btnProjectLeader.MouseLeave += MouseLeaveStatus;

            txtProjectLeader.Tag = "מתכנן מוביל בפרויקט";
            txtProjectLeader.MouseEnter += MouseEnterStatus; txtProjectLeader.MouseLeave += MouseLeaveStatus;

            txtProjectID.Tag = "הכנס מספר פרויקט";
            txtProjectID.MouseEnter += MouseEnterStatus; txtProjectID.MouseLeave += MouseLeaveStatus;

            txtProjectName.Tag = "שם הפרויקט";
            txtProjectName.MouseEnter += MouseEnterStatus; txtProjectName.MouseLeave += MouseLeaveStatus;

            txtSubject.Tag = "נושא ההודעה";
            txtSubject.MouseEnter += MouseEnterStatus; txtSubject.MouseLeave += MouseLeaveStatus;

            txtFullPath.Tag = "נתיב מלא";
            txtFullPath.MouseEnter += MouseEnterStatus; txtFullPath.MouseLeave += MouseLeaveStatus;

            cmbSaveLocation.Tag = "בחר מיקום שמירה";
            cmbSaveLocation.MouseEnter += MouseEnterStatus; cmbSaveLocation.MouseLeave += MouseLeaveStatus;

            rtxtNotes.Tag = "הערות למייל";
            rtxtNotes.MouseEnter += MouseEnterStatus; rtxtNotes.MouseLeave += MouseLeaveStatus;

            rtxtProjectNotes.Tag = "הערות בפרויקט";
            rtxtProjectNotes.MouseEnter += MouseEnterStatus; rtxtProjectNotes.MouseLeave += MouseLeaveStatus;

            chkbSendNote.Tag = "שלח ההערה לראש הפרויקט";
            chkbSendNote.MouseEnter += MouseEnterStatus; chkbSendNote.MouseLeave += MouseLeaveStatus;

            chkbSelectAllAttachments.Tag = "בחר/הסר כל הקבצים";
            chkbSelectAllAttachments.MouseEnter += MouseEnterStatus; chkbSelectAllAttachments.MouseLeave += MouseLeaveStatus;

            chbOpenPDF.Tag = "פתח PDF לאחר שמירה";
            chbOpenPDF.MouseEnter += MouseEnterStatus; chbOpenPDF.MouseLeave += MouseLeaveStatus;

            tvFolders.Tag = "עץ תיקיות פרויקט";
            tvFolders.MouseEnter += MouseEnterStatus; tvFolders.MouseLeave += MouseLeaveStatus;

            dgvAttachments.Tag = "קבצים מצורפים";
            dgvAttachments.MouseEnter += MouseEnterStatus; dgvAttachments.MouseLeave += MouseLeaveStatus;

            dgvEmployees.Tag = "עובדים בפרויקט";
            dgvEmployees.MouseEnter += MouseEnterStatus; dgvEmployees.MouseLeave += MouseLeaveStatus;

            tabNotes.Tag = "כרטיסיות הערות";
            tabNotes.MouseEnter += MouseEnterStatus; tabNotes.MouseLeave += MouseLeaveStatus;

            tabFilesFolders.Tag = "קבצים ותיקיות";
            tabFilesFolders.MouseEnter += MouseEnterStatus; tabFilesFolders.MouseLeave += MouseLeaveStatus;
        }

        private void dgvAttachments_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (dgvAttachments.CurrentCell != null)
            {
                XMessageBox.Show(
                    dgvAttachments.CurrentCell.Value?.ToString() ?? string.Empty,
                    "פרטי קובץ",
                    XMessageBoxButtons.OK,
                    XMessageBoxIcon.Information,
                    XMessageAlignment.Right,
                    XMessageLanguage.Hebrew);
            }
        }

        private void btnStyle_Click(object sender, EventArgs e)
        {
            if (dlgFont.ShowDialog() == DialogResult.OK)
                rtxtNotes.SelectionFont = dlgFont.Font;
        }

        private void btnPhoneBook_Click(object sender, EventArgs e)
        {
            using (var frmContacts = new FormContacts(this))
            {
                frmContacts.ShowDialog();
            }
        }

        private void chkbSelectAllAttachments_CheckedChanged(object sender, EventArgs e)
        {
            chkbSelectAllAttachments.Text = chkbSelectAllAttachments.Checked ? "הסר הכל" : "בחר הכל";
            if (dgvAttachments.RowCount != 0)
            {
                dgvAttachments.SuspendLayout();
                try
                {
                    dgvAttachments.BeginEdit(false);
                    foreach (DataGridViewRow row in dgvAttachments.Rows)
                        row.Cells[1].Value = chkbSelectAllAttachments.Checked;
                    dgvAttachments.EndEdit();
                }
                finally
                {
                    dgvAttachments.ResumeLayout(false);
                }
            }
        }

        private void btnCopyNotesToMail_Click(object sender, EventArgs e)
        {
            rtxtNotes.Text += "\n " + rtxtProjectNotes.Text;
        }

        private void btnCopyNotesToProject_Click(object sender, EventArgs e)
        {
            rtxtProjectNotes.Text += "\n " + rtxtNotes.Text;
        }

        private void tvFolders_BeforeExpand(object sender, TreeViewCancelEventArgs e)
        {
            if (_isDoubleClick && e.Action == TreeViewAction.Expand)
            {
                e.Cancel = true;
                return;
            }

            if (e.Node.Nodes.Count == 1 && e.Node.Nodes[0].Text == "...")
            {
                try
                {
                    string basePath = e.Node.Tag as string;
                    if (!string.IsNullOrEmpty(basePath) && Directory.Exists(basePath))
                    {
                        e.Node.Nodes.Clear();
                        var nodes = TreeHelpers.GetFolderNodes(basePath, expanded: false);
                        foreach (var n in nodes)
                            e.Node.Nodes.Add(n);
                    }
                }
                catch (Exception ex)
                {
                    XMessageBox.Show(
                        $"שגיאה בטעינת תיקיות: {ex.Message}",
                        "SaveAsPDF:tvFolders_BeforeExpand",
                        XMessageBoxButtons.OK,
                        XMessageBoxIcon.Error,
                        XMessageAlignment.Right,
                        XMessageLanguage.Hebrew);
                }
            }
        }

        private void tvFolders_BeforeCollapse(object sender, TreeViewCancelEventArgs e)
        {
            if (_isDoubleClick && e.Action == TreeViewAction.Collapse)
                e.Cancel = true;
        }

        private void tvFolders_MouseDown(object sender, MouseEventArgs e)
        {
            _isDoubleClick = e.Clicks > 1;
        }

        private void tvFolders_NodeMouseClick(object sender, TreeNodeMouseClickEventArgs e)
        {
            if (e.Node.Tag is string basePath)
            {
                cmbSaveLocation.Select();
                cmbSaveLocation.SelectedItem = null;
                cmbSaveLocation.SelectedText = basePath;
            }
        }

        private void tvFolders_NodeMouseDoubleClick(object sender, TreeNodeMouseClickEventArgs e)
        {
            if (e.Node.Tag is string basePath)
            {
                System.Diagnostics.Process.Start("explorer.exe", basePath);
                cmbSaveLocation.SelectedText = basePath;
            }
        }

        private void dgvEmployees_CurrentCellDirtyStateChanged(object sender, EventArgs e)
        {
            if (dgvEmployees.IsCurrentCellDirty && dgvEmployees.CurrentCell is DataGridViewCheckBoxCell)
            {
                dgvEmployees.CommitEdit(DataGridViewDataErrorContexts.Commit);
            }
        }

        private void dgvEmployees_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            // Kept minimal: you can extend with leader logic if needed
        }

        private void cmbSaveLocation_SelectedValueChanged(object sender, EventArgs e)
        {
            // Stub handler
        }

        private void cmbSaveLocation_TextUpdate(object sender, EventArgs e)
        {
            // Stub handler
        }

        private void btnRemoveEmployee_Click(object sender, EventArgs e)
        {
            // Stub handler
        }
    }
}