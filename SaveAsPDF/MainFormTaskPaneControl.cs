// Task-pane version of FormMain with full right-to-left layout and core logic
using Microsoft.Office.Interop.Outlook;
using SaveAsPDF.Helpers;
using SaveAsPDF.Models;
using SaveAsPDF.Properties;
using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;
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
        private readonly ExplorerAddressBar cmbSaveLocation = new ExplorerAddressBar();
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
        private TabControl tabNotes = new TabControl();
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

            // Apply theme-aware colors
            ApplySystemColors();

            BuildLayout();
            ConfigureEmployeeDataGrid();
            ConfigureAttachmentsDataGrid();

            KeyDown += MainFormTaskPaneControl_KeyDown;

            btnProjectLeader.Click += btnProjectLeader_Click;
            btnSettings.CausesValidation = false;
            btnCancel.CausesValidation = false;

            Load += MainFormTaskPaneControl_Load;
        }

        private void ApplySystemColors()
        {
            // Set base control colors to match system theme
            BackColor = SystemColors.Control;
            ForeColor = SystemColors.ControlText;
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
            cmbSaveLocation.Path = string.Empty;
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

            // Load project data from XML
            LoadProjectData();

            // Load employee data from XML
            LoadEmployeeData();

            // Update the UI with loaded data
            UpdateUI();
        }

        private void LoadProjectData()
        {
            if (File.Exists(settingsModel.XmlProjectFile))
            {
                try
                {
                    _projectModel = XmlFileHelper.XmlProjectFileToModel(settingsModel.XmlProjectFile);
                    if (_projectModel != null)
                    {
                        txtProjectName.Text = _projectModel.ProjectName;
                        rtxtProjectNotes.Text = _projectModel.ProjectNotes;
                    }
                }
                catch (Exception ex)
                {
                    XMessageBox.Show(
                        $"שגיאה בטעינת נתוני הפרויקט: {ex.Message}",
                        "SaveAsPDF:LoadProjectData",
                        XMessageBoxButtons.OK,
                        XMessageBoxIcon.Warning,
                        XMessageAlignment.Right,
                        XMessageLanguage.Hebrew);
                    SetDefaultProjectModel();
                }
            }
            else
            {
                SetDefaultProjectModel();
            }
        }

        private void SetDefaultProjectModel()
        {
            _projectModel = new ProjectModel
            {
                ProjectName = "פרויקט ברירת מחדל",
                ProjectNumber = txtProjectID.Text,
                NoteToProjectLeader = false,
                DefaultSaveFolder = settingsModel.DefaultSavePath,
                ProjectNotes = "הערות ברירת מחדל",
                LastSavePath = settingsModel.DefaultSavePath
            };

            if (!string.IsNullOrEmpty(settingsModel.XmlProjectFile))
            {
                settingsModel.XmlProjectFile.ProjectModelToXmlFile(_projectModel);
            }

            txtProjectName.Text = _projectModel.ProjectName;
            rtxtProjectNotes.Text = _projectModel.ProjectNotes;
        }

        private void LoadEmployeeData()
        {
            _employeesBindingList.Clear();

            if (File.Exists(settingsModel.XmlEmployeesFile))
            {
                var loaded = settingsModel.XmlEmployeesFile.XmlEmployeesFileToModel();
                if (loaded != null)
                {
                    foreach (var em in loaded)
                    {
                        _employeesBindingList.Add(em);
                    }
                }

                // Update project leader textbox from the first employee marked as leader
                var leader = _employeesBindingList.FirstOrDefault(e => e.IsLeader);
                if (leader != null)
                {
                    var first = leader.FirstName ?? string.Empty;
                    var last = leader.LastName ?? string.Empty;
                    string fullName = ($"{first} {last}").Trim();
                    if (string.IsNullOrWhiteSpace(fullName))
                    {
                        fullName = leader.EmailAddress ?? string.Empty;
                    }
                    txtProjectLeader.Text = fullName;
                }
                else
                {
                    txtProjectLeader.Clear();
                }
            }
        }

        private void UpdateUI()
        {
            try
            {
                string targetPath = settingsModel.DefaultSavePath;
                if (string.IsNullOrEmpty(targetPath) && settingsModel.ProjectRootFolder != null && settingsModel.ProjectRootFolder.Exists)
                {
                    targetPath = settingsModel.ProjectRootFolder.FullName;
                }

                if (!string.IsNullOrEmpty(targetPath))
                {
                    cmbSaveLocation.Path = targetPath;
                    txtFullPath.Text = PathBreadcrumbHelper.FormatPathAsBreadcrumb(targetPath);
                }

                // Update the tree view with the project root
                tvFolders.Nodes.Clear();
                if (settingsModel.ProjectRootFolder != null && settingsModel.ProjectRootFolder.Exists)
                {
                    var rootDir = settingsModel.ProjectRootFolder;
                    var rootNode = new TreeNode(rootDir.Name) { Tag = rootDir.FullName };
                    rootNode.Nodes.Add("...");
                    tvFolders.Nodes.Add(rootNode);
                    tvFolders.SelectedNode = tvFolders.Nodes[0];

                    txtFullPath.Text = PathBreadcrumbHelper.FormatPathAsBreadcrumb(rootDir.FullName);
                }
                else
                {
                    XMessageBox.Show(
                        "תיקיית השורש של הפרויקט אינה קיימת.",
                        "שגיאה",
                        XMessageBoxButtons.OK,
                        XMessageBoxIcon.Warning,
                        XMessageAlignment.Right,
                        XMessageLanguage.Hebrew);
                }
            }
            catch (Exception ex)
            {
                XMessageBox.Show(
                    $"אירעה שגיאהבעת עדכון הממשק: {ex.Message}",
                    "שגיאה",
                    XMessageBoxButtons.OK,
                    XMessageBoxIcon.Error,
                    XMessageAlignment.Right,
                    XMessageLanguage.Hebrew);
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
                RightToLeft = RightToLeft.Yes,
                BackColor = SystemColors.Control,
                ForeColor = SystemColors.ControlText
            };

            var topTable = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 2,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                RightToLeft = RightToLeft.Yes,
                BackColor = SystemColors.Control
            };
            topTable.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 30));
            topTable.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 70));

            // Project ID
            var lblProjectID = new Label
            {
                Text = "מספר פרויקט",
                AutoSize = true,
                TextAlign = ContentAlignment.MiddleLeft,
                Dock = DockStyle.Fill,
                BackColor = SystemColors.Control,
                ForeColor = SystemColors.ControlText
            };
            txtProjectID.Dock = DockStyle.Fill;
            txtProjectID.BackColor = SystemColors.Window;
            txtProjectID.ForeColor = SystemColors.WindowText;
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
                Dock = DockStyle.Fill,
                BackColor = SystemColors.Control,
                ForeColor = SystemColors.ControlText
            };
            txtProjectName.Dock = DockStyle.Fill;
            txtProjectName.BackColor = SystemColors.Window;
            txtProjectName.ForeColor = SystemColors.WindowText;
            topTable.Controls.Add(lblProjectName, 0, 1);
            topTable.Controls.Add(txtProjectName, 1, 1);

            // Project leader
            var lblLeader = new Label
            {
                Text = "מתכנן מוביל",
                AutoSize = true,
                TextAlign = ContentAlignment.MiddleLeft,
                Dock = DockStyle.Fill,
                BackColor = SystemColors.Control,
                ForeColor = SystemColors.ControlText
            };
            var pnlLeader = new FlowLayoutPanel
            {
                FlowDirection = FlowDirection.LeftToRight,
                Dock = DockStyle.Fill,
                AutoSize = true,
                BackColor = SystemColors.Control
            };
            txtProjectLeader.Width = 160;
            txtProjectLeader.BackColor = SystemColors.Window;
            txtProjectLeader.ForeColor = SystemColors.WindowText;
            btnProjectLeader.Text = "...";
            btnProjectLeader.Width = 30;
            btnProjectLeader.BackColor = SystemColors.Control;
            btnProjectLeader.ForeColor = SystemColors.ControlText;
            btnProjectLeader.UseVisualStyleBackColor = true;
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
                Dock = DockStyle.Fill,
                BackColor = SystemColors.Control,
                ForeColor = SystemColors.ControlText
            };
            txtSubject.Dock = DockStyle.Fill;
            txtSubject.BackColor = SystemColors.Window;
            txtSubject.ForeColor = SystemColors.WindowText;
            topTable.Controls.Add(lblSubject, 0, 3);
            topTable.Controls.Add(txtSubject, 1, 3);

            // Full path
            var lblFullPath = new Label
            {
                Text = "נתיב מלא",
                AutoSize = true,
                TextAlign = ContentAlignment.MiddleLeft,
                Dock = DockStyle.Fill,
                BackColor = SystemColors.Control,
                ForeColor = SystemColors.ControlText
            };
            txtFullPath.Dock = DockStyle.Fill;
            txtFullPath.ReadOnly = true;
            txtFullPath.BackColor = SystemColors.Window; // Change from Control to Window to look like address bar
            txtFullPath.ForeColor = SystemColors.WindowText;
            txtFullPath.BorderStyle = BorderStyle.FixedSingle; // Add border like an address bar
            txtFullPath.Cursor = Cursors.Arrow; // Change cursor since it's read-only
            txtFullPath.RightToLeft = RightToLeft.No; // Align content to the left
            txtFullPath.TextAlign = HorizontalAlignment.Left;
            topTable.Controls.Add(lblFullPath, 0, 4);
            topTable.Controls.Add(txtFullPath, 1, 4);

            // Save location combo + folder selector
            var lblSaveLocation = new Label
            {
                Text = "מיקום שמירה",
                AutoSize = true,
                TextAlign = ContentAlignment.MiddleLeft,
                Dock = DockStyle.Fill,
                BackColor = SystemColors.Control,
                ForeColor = SystemColors.ControlText
            };
            var pnlSaveLocation = new FlowLayoutPanel
            {
                FlowDirection = FlowDirection.LeftToRight,
                Dock = DockStyle.Fill,
                AutoSize = true,
                BackColor = SystemColors.Control,
                RightToLeft = RightToLeft.No
            };
            cmbSaveLocation.Width = 280;
            cmbSaveLocation.BackColor = SystemColors.Window;
            cmbSaveLocation.ForeColor = SystemColors.WindowText;
            cmbSaveLocation.RightToLeft = RightToLeft.No;
            cmbSaveLocation.PathConfirmed += CmbSaveLocation_PathConfirmed;
            btnFolders.Text = "בחר תיקייה";
            btnFolders.BackColor = SystemColors.Control;
            btnFolders.ForeColor = SystemColors.ControlText;
            btnFolders.UseVisualStyleBackColor = true;
            btnFolders.Click += btnFolders_Click;
            pnlSaveLocation.Controls.Add(cmbSaveLocation);
            pnlSaveLocation.Controls.Add(btnFolders);
            topTable.Controls.Add(lblSaveLocation, 0, 5);
            topTable.Controls.Add(pnlSaveLocation, 1, 5);

            topGroup.Controls.Add(topTable);

            // --- Middle: Single unified TabControl ---
            var mainTabControl = new TabControl
            {
                Dock = DockStyle.Fill,
                RightToLeft = RightToLeft.Yes,
                Alignment = TabAlignment.Right
            };

            // Notes Tab with Sub-TabControl
            var tabNotes = new TabPage("הערות")
            {
                RightToLeft = RightToLeft.Yes,
                BackColor = SystemColors.Control,
                UseVisualStyleBackColor = true
            };
            var subTabNotes = new TabControl
            {
                Dock = DockStyle.Fill,
                RightToLeft = RightToLeft.Yes,
                Alignment = TabAlignment.Right
            };

            // Mail Notes Sub-Tab
            var tabMailNotes = new TabPage("הערות למייל")
            {
                RightToLeft = RightToLeft.Yes,
                BackColor = SystemColors.Control,
                UseVisualStyleBackColor = true
            };
            rtxtNotes.Dock = DockStyle.Fill;
            rtxtNotes.BackColor = SystemColors.Window;
            rtxtNotes.ForeColor = SystemColors.WindowText;
            tabMailNotes.Controls.Add(rtxtNotes);

            // Project Notes Sub-Tab
            var tabProjectNotes = new TabPage("הערות בפרויקט")
            {
                RightToLeft = RightToLeft.Yes,
                BackColor = SystemColors.Control,
                UseVisualStyleBackColor = true
            };
            rtxtProjectNotes.Dock = DockStyle.Fill;
            rtxtProjectNotes.BackColor = SystemColors.Window;
            rtxtProjectNotes.ForeColor = SystemColors.WindowText;
            tabProjectNotes.Controls.Add(rtxtProjectNotes);

            // Add sub-tabs to sub-TabControl
            subTabNotes.TabPages.Add(tabMailNotes);
            subTabNotes.TabPages.Add(tabProjectNotes);

            tabNotes.Controls.Add(subTabNotes);

            // Attachments Tab
            var tabAttachments = new TabPage("קבצים מצורפים")
            {
                RightToLeft = RightToLeft.Yes,
                BackColor = SystemColors.Control,
                UseVisualStyleBackColor = true
            };
            var attachmentsTable = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                RowCount = 2,
                ColumnCount = 1,
                BackColor = SystemColors.Control
            };
            attachmentsTable.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            attachmentsTable.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            chkbSelectAllAttachments.Text = "הסר הכל";
            chkbSelectAllAttachments.Checked = true;
            chkbSelectAllAttachments.BackColor = SystemColors.Control;
            chkbSelectAllAttachments.ForeColor = SystemColors.ControlText;
            chkbSelectAllAttachments.CheckedChanged += chkbSelectAllAttachments_CheckedChanged;
            attachmentsTable.Controls.Add(chkbSelectAllAttachments, 0, 0);
            dgvAttachments.Dock = DockStyle.Fill;
            dgvAttachments.CellDoubleClick += dgvAttachments_CellDoubleClick;
            attachmentsTable.Controls.Add(dgvAttachments, 0, 1);
            tabAttachments.Controls.Add(attachmentsTable);

            // Employees Tab
            var tabEmployees = new TabPage("עובדים בפרויקט")
            {
                RightToLeft = RightToLeft.Yes,
                BackColor = SystemColors.Control,
                UseVisualStyleBackColor = true
            };
            var employeesTable = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                RowCount = 2,
                ColumnCount = 1,
                BackColor = SystemColors.Control
            };
            employeesTable.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            employeesTable.RowStyles.Add(new RowStyle(SizeType.Percent, 100));

            var pnlEmpButtons = new FlowLayoutPanel
            {
                FlowDirection = FlowDirection.RightToLeft,
                Dock = DockStyle.Fill,
                AutoSize = true,
                BackColor = SystemColors.Control
            };
            btnPhoneBook.Text = "ספר טלפונים";
            btnPhoneBook.BackColor = SystemColors.Control;
            btnPhoneBook.ForeColor = SystemColors.ControlText;
            btnPhoneBook.UseVisualStyleBackColor = true;
            btnPhoneBook.Click += btnPhoneBook_Click;
            btnRemoveEmployee.Text = "הסר עובד";
            btnRemoveEmployee.BackColor = SystemColors.Control;
            btnRemoveEmployee.ForeColor = SystemColors.ControlText;
            btnRemoveEmployee.UseVisualStyleBackColor = true;
            btnRemoveEmployee.Click += btnRemoveEmployee_Click;
            pnlEmpButtons.Controls.Add(btnPhoneBook);
            pnlEmpButtons.Controls.Add(btnRemoveEmployee);

            employeesTable.Controls.Add(pnlEmpButtons, 0, 0);
            dgvEmployees.Dock = DockStyle.Fill;
            dgvEmployees.CurrentCellDirtyStateChanged += dgvEmployees_CurrentCellDirtyStateChanged;
            dgvEmployees.CellValueChanged += dgvEmployees_CellValueChanged;
            employeesTable.Controls.Add(dgvEmployees, 0, 1);
            tabEmployees.Controls.Add(employeesTable);

            // Folders Tab
            var tabFolders = new TabPage("עץ תיקיות פרויקט")
            {
                RightToLeft = RightToLeft.Yes,
                BackColor = SystemColors.Control,
                UseVisualStyleBackColor = true
            };
            tvFolders.Dock = DockStyle.Fill;
            tvFolders.BackColor = SystemColors.Window;
            tvFolders.ForeColor = SystemColors.WindowText;
            tvFolders.PathSeparator = "\\";
            tvFolders.BeforeExpand += tvFolders_BeforeExpand;
            tvFolders.BeforeCollapse += tvFolders_BeforeCollapse;
            tvFolders.MouseDown += tvFolders_MouseDown;
            tvFolders.NodeMouseClick += tvFolders_NodeMouseClick;
            tvFolders.NodeMouseDoubleClick += tvFolders_NodeMouseDoubleClick;
            tabFolders.Controls.Add(tvFolders);

            // Add all tabs to main TabControl
            mainTabControl.TabPages.Add(tabAttachments);
            mainTabControl.TabPages.Add(tabEmployees);
            mainTabControl.TabPages.Add(tabFolders);
            mainTabControl.TabPages.Add(tabNotes);

            // Wrapper with buttons for notes tabs
            var middlePanel = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                RowCount = 2,
                ColumnCount = 1,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                BackColor = SystemColors.Control
            };
            middlePanel.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            middlePanel.RowStyles.Add(new RowStyle(SizeType.AutoSize));

            var pnlNotesButtons = new FlowLayoutPanel
            {
                FlowDirection = FlowDirection.RightToLeft,
                Dock = DockStyle.Fill,
                AutoSize = true,
                BackColor = SystemColors.Control
            };
            btnCopyNotesToMail.Text = "העתק לפרויקט ← מייל";
            btnCopyNotesToMail.BackColor = SystemColors.Control;
            btnCopyNotesToMail.ForeColor = SystemColors.ControlText;
            btnCopyNotesToMail.UseVisualStyleBackColor = true;
            btnCopyNotesToMail.Click += btnCopyNotesToMail_Click;
            btnCopyNotesToProject.Text = "העתק ממייל → פרויקט";
            btnCopyNotesToProject.BackColor = SystemColors.Control;
            btnCopyNotesToProject.ForeColor = SystemColors.ControlText;
            btnCopyNotesToProject.UseVisualStyleBackColor = true;
            btnCopyNotesToProject.Click += btnCopyNotesToProject_Click;
            btnStyle.Text = "גופן";
            btnStyle.BackColor = SystemColors.Control;
            btnStyle.ForeColor = SystemColors.ControlText;
            btnStyle.UseVisualStyleBackColor = true;
            btnStyle.Click += btnStyle_Click;
            pnlNotesButtons.Controls.Add(btnCopyNotesToMail);
            pnlNotesButtons.Controls.Add(btnCopyNotesToProject);
            pnlNotesButtons.Controls.Add(btnStyle);

            middlePanel.Controls.Add(mainTabControl, 0, 0);
            middlePanel.Controls.Add(pnlNotesButtons, 0, 1);

            // Store reference to main tab control
            this.tabNotes = mainTabControl;

            // --- Bottom buttons + status ---
            var bottomPanel = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 2,
                AutoSize = true,
                BackColor = SystemColors.Control
            };
            bottomPanel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            bottomPanel.RowStyles.Add(new RowStyle(SizeType.AutoSize));

            var pnlButtons = new FlowLayoutPanel
            {
                FlowDirection = FlowDirection.RightToLeft,
                Dock = DockStyle.Fill,
                AutoSize = true,
                BackColor = SystemColors.Control
            };
            btnOK.Text = "שמור ל-PDF";
            btnOK.BackColor = SystemColors.Control;
            btnOK.ForeColor = SystemColors.ControlText;
            btnOK.UseVisualStyleBackColor = true;
            btnOK.Click += btnOK_Click;
            btnCancel.Text = "בטل";
            btnCancel.BackColor = SystemColors.Control;
            btnCancel.ForeColor = SystemColors.ControlText;
            btnCancel.UseVisualStyleBackColor = true;
            btnCancel.Click += btnCancel_Click;
            btnSettings.Text = "הגדרות";
            btnSettings.BackColor = SystemColors.Control;
            btnSettings.ForeColor = SystemColors.ControlText;
            btnSettings.UseVisualStyleBackColor = true;
            btnSettings.Click += BtnSettings_Click;
            btnNewProject.Text = "פרויקט חדש";
            btnNewProject.BackColor = SystemColors.Control;
            btnNewProject.ForeColor = SystemColors.ControlText;
            btnNewProject.UseVisualStyleBackColor = true;
            btnNewProject.Click += btnNewProject_Click;
            chbOpenPDF.Text = "פתח PDF לאחר שמירה";
            chbOpenPDF.BackColor = SystemColors.Control;
            chbOpenPDF.ForeColor = SystemColors.ControlText;
            chbOpenPDF.CheckedChanged += chbOpenPDF_CheckedChanged;

            pnlButtons.Controls.Add(btnOK);
            pnlButtons.Controls.Add(btnCancel);
            pnlButtons.Controls.Add(btnSettings);
            pnlButtons.Controls.Add(btnNewProject);
            pnlButtons.Controls.Add(chbOpenPDF);

            statusStrip.Items.Add(tsslStatus);
            statusStrip.RightToLeft = RightToLeft.Yes;
            statusStrip.BackColor = SystemColors.Control;
            statusStrip.ForeColor = SystemColors.ControlText;

            bottomPanel.Controls.Add(pnlButtons, 0, 0);
            bottomPanel.Controls.Add(statusStrip, 0, 1);

            main.Controls.Add(topGroup, 0, 0);
            main.Controls.Add(middlePanel, 0, 1);
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

                // Update both save location and full path with breadcrumb format
                cmbSaveLocation.Path = dialog.ResultPath;
                txtFullPath.Text = PathBreadcrumbHelper.FormatPathAsBreadcrumb(dialog.ResultPath);
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

                // Process mail item attachments if available
                if (_mailItem != null)
                {
                    ProcessMailItem(_mailItem);
                }
            }
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            // Get actual path from address bar
            string sPath = cmbSaveLocation.Path;

            if (string.IsNullOrEmpty(sPath))
            {
                var dialog = new FolderPicker { InputPath = settingsModel.RootDrive };
                if (dialog.ShowDialog(Handle) == true)
                {
                    sPath = dialog.ResultPath;
                    cmbSaveLocation.Path = sPath;
                }
            }

            if (!string.IsNullOrEmpty(sPath))
            {
                // Validate path for illegal characters
                try
                {
                    // This will throw if path contains illegal characters
                    Path.GetFullPath(sPath);
                }
                catch (ArgumentException)
                {
                    XMessageBox.Show(
                    "הנתיב מכיל תווים לא חוקיים. אנא בחר נתיב תקין.",
                    "שגיאה",
                    XMessageBoxButtons.OK,
                    XMessageBoxIcon.Error,
                    XMessageAlignment.Right,
                    XMessageLanguage.Hebrew);
                    return;
                }
                catch (Exception ex)
                {
                    XMessageBox.Show(
                    $"נתיב לא תקין: {ex.Message}",
                    "שגיאה",
                    XMessageBoxButtons.OK,
                    XMessageBoxIcon.Error,
                    XMessageAlignment.Right,
                    XMessageLanguage.Hebrew);
                    return;
                }

                var directory = new DirectoryInfo(sPath);
                if (!directory.Exists)
                    FileFoldersHelper.CreateDirectory(directory.FullName);
            }
            else
            {
                XMessageBox.Show(
                "יש לבחור או לציין מיקום שמירה תקין.",
                "שגיאה",
                XMessageBoxButtons.OK,
                XMessageBoxIcon.Error,
                XMessageAlignment.Right,
                XMessageLanguage.Hebrew);
                return;
            }

            // Save selected attachments to the chosen folder
            List<AttachmentsModel> savedAttachmentsModels = new List<AttachmentsModel>();
            try
            {
                var saved = _mailItem.SaveAttachments(dgvAttachments, sPath, overWrite: false);
                int idx = 0;
                foreach (var entry in saved)
                {
                    var parts = entry.Split(new[] { '|' }, 2);
                    var fname = parts.Length > 0 ? parts[0] : string.Empty;
                    var fsize = parts.Length > 1 ? parts[1] : string.Empty;
                    savedAttachmentsModels.Add(new AttachmentsModel
                    {
                        attachmentId = idx++,
                        isChecked = true,
                        fileName = fname,
                        fileSize = fsize
                    });
                }
            }
            catch (Exception ex)
            {
                XMessageBox.Show(
                    $"שגיאה בשמירת קבצים מצורפים: {ex.Message}",
                    "SaveAsPDF",
                    XMessageBoxButtons.OK,
                    XMessageBoxIcon.Warning,
                    XMessageAlignment.Right,
                    XMessageLanguage.Hebrew
                );
            }

            // Generate HTML to file
            string sanitizedProjectName = txtProjectName.Text.SafeFolderName();
            string timeStamp = DateTime.Now.ToString("yyyyMMddHHmmss");
            string htmlFileName = $"{sanitizedProjectName}_{timeStamp}.html";
            string htmlFilePath = Path.Combine(sPath, htmlFileName);

            try
            {
                HtmlHelper.GenerateHtmlToFile(
                htmlFilePath,
                sPath,
                _employeesBindingList.ToList(),
                (savedAttachmentsModels.Count > 0) ? savedAttachmentsModels : attachmentsModels,
                txtProjectName.Text,
                txtProjectID.Text,
                rtxtNotes.Text,
                Environment.UserName,
                (_mailItem != null && !string.IsNullOrWhiteSpace(_mailItem.Subject)) ? _mailItem.Subject : txtSubject.Text
                );

                try
                {
                    string htmlContent = File.ReadAllText(htmlFilePath);
                    _mailItem.HTMLBody = htmlContent + _mailItem.HTMLBody;
                }
                catch (Exception ex)
                {
                    XMessageBox.Show(
                    $"שגיאה בטעינת קובץ HTML שנוצר: {ex.Message}",
                    "SaveAsPDF",
                    XMessageBoxButtons.OK,
                    XMessageBoxIcon.Warning,
                    XMessageAlignment.Right,
                    XMessageLanguage.Hebrew);
                }
                finally
                {
                    try { if (File.Exists(htmlFilePath)) File.Delete(htmlFilePath); } catch { }
                }
            }
            catch (Exception ex)
            {
                XMessageBox.Show(
                $"שגיאה ביצירת קובץ ה-HTML: {ex.Message}",
                "SaveAsPDF",
                XMessageBoxButtons.OK,
                XMessageBoxIcon.Error,
                XMessageAlignment.Right,
                XMessageLanguage.Hebrew
                );
                return;
            }

            _mailItem.SaveToPDF(sPath);
            _mailItem.Save();

            // If requested, forward the current email to the project leader
            if (chkbSendNote.Checked)
            {
                try
                {
                    var leader = _employeesBindingList.FirstOrDefault(emp => emp.IsLeader && !string.IsNullOrWhiteSpace(emp.EmailAddress));
                    var leaderEmail = leader?.EmailAddress;
                    if (!string.IsNullOrWhiteSpace(leaderEmail))
                    {
                        var fwd = _mailItem.Forward();
                        fwd.To = leaderEmail;
                    }
                    else
                    {
                        XMessageBox.Show(
                        "לא encontrado אימייל של ראש הפרויקט. יש לבחור ראש פרויקט או לעדכן את כתובת האימייל שלו.",
                        "SaveAsPDF",
                        XMessageBoxButtons.OK,
                        XMessageBoxIcon.Warning,
                        XMessageAlignment.Right,
                        XMessageLanguage.Hebrew
                        );
                    }
                }
                catch (Exception ex)
                {
                    XMessageBox.Show(
                    $"שגיאה בשליחת הודעת העברה לראש הפרויקט: {ex.Message}",
                    "SaveAsPDF",
                    XMessageBoxButtons.OK,
                    XMessageBoxIcon.Error,
                    XMessageAlignment.Right,
                    XMessageLanguage.Hebrew
                    );
                }
            }

            // Clear mail item (task pane doesn't close like form)
            if (_mailItem != null)
            {
                _mailItem = null;
                ClearMailRelatedUi();
            }

            if (chbOpenPDF.Checked)
            {
                // Sanitize subject text for use in filename
                string sanitizedSubject = txtSubject.Text.SafeFolderName();
                string pdfFilePath = Path.Combine(sPath, $"{sanitizedSubject}.pdf");
                if (File.Exists(pdfFilePath))
                {
                    System.Diagnostics.Process.Start(pdfFilePath);
                }
                else
                {
                    XMessageBox.Show(
                    "קובץ ה-PDF לא נמצא.",
                    "שגיאה",
                    XMessageBoxButtons.OK,
                    XMessageBoxIcon.Error,
                    XMessageAlignment.Right,
                    XMessageLanguage.Hebrew);
                }
            }
        }

        private void btnRemoveEmployee_Click(object sender, EventArgs e)
        {
            // Stub handler
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
            if (mailItem == null)
                return;

            txtSubject.Text = mailItem.Subject;

            var attachments = mailItem.GetAttachmentsFromEmail();
            int i = 0;
            attachmentsModels.Clear();

            foreach (var attachment in attachments)
            {
                if (attachment != null)
                {
                    attachmentsModels.Add(new AttachmentsModel
                    {
                        attachmentId = i++,
                        isChecked = true,
                        fileName = attachment.FileName,
                        fileSize = attachment.Size.BytesToString()
                    });
                }
            }

            dgvAttachments.DataSource = null;
            dgvAttachments.DataSource = attachmentsModels;

            if (dgvAttachments.Columns.Count > 0)
            {
                dgvAttachments.Columns[0].Visible = false;

                if (dgvAttachments.Columns.Count > 1)
                {
                    dgvAttachments.Columns[1].HeaderText = "V";
                    dgvAttachments.Columns[1].ReadOnly = false;
                }

                if (dgvAttachments.Columns.Count > 2)
                {
                    dgvAttachments.Columns[2].HeaderText = "שם קובץ";
                    dgvAttachments.Columns[2].ReadOnly = true;
                }

                if (dgvAttachments.Columns.Count > 3)
                {
                    dgvAttachments.Columns[3].HeaderText = "גודל";
                    dgvAttachments.Columns[3].ReadOnly = true;
                }
            }
        }

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
                cmbSaveLocation.Path = basePath;
                txtFullPath.Text = PathBreadcrumbHelper.FormatPathAsBreadcrumb(basePath);
            }
        }

        private void tvFolders_NodeMouseDoubleClick(object sender, TreeNodeMouseClickEventArgs e)
        {
            if (e.Node.Tag is string basePath)
            {
                System.Diagnostics.Process.Start("explorer.exe", basePath);
                cmbSaveLocation.Path = basePath;
                txtFullPath.Text = PathBreadcrumbHelper.FormatPathAsBreadcrumb(basePath);
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

            dgvEmployees.BackgroundColor = SystemColors.Window;
            dgvEmployees.ForeColor = SystemColors.WindowText;
            dgvEmployees.DefaultCellStyle.BackColor = SystemColors.Window;
            dgvEmployees.DefaultCellStyle.ForeColor = SystemColors.WindowText;
            dgvEmployees.DefaultCellStyle.SelectionBackColor = SystemColors.Highlight;
            dgvEmployees.DefaultCellStyle.SelectionForeColor = SystemColors.HighlightText;
            dgvEmployees.AlternatingRowsDefaultCellStyle.BackColor = SystemColors.Control;
            dgvEmployees.AlternatingRowsDefaultCellStyle.ForeColor = SystemColors.ControlText;
            dgvEmployees.ColumnHeadersDefaultCellStyle.BackColor = SystemColors.Control;
            dgvEmployees.ColumnHeadersDefaultCellStyle.ForeColor = SystemColors.ControlText;
            dgvEmployees.EnableHeadersVisualStyles = false;
        }

        private void ConfigureAttachmentsDataGrid()
        {
            dgvAttachments.AutoGenerateColumns = false;
            dgvAttachments.Columns.Clear();

            dgvAttachments.Columns.AddRange(new DataGridViewColumn[]
            {
                new DataGridViewTextBoxColumn
                {
                    Name = "attachmentId",
                    DataPropertyName = "attachmentId",
                    Visible = false
                },
                new DataGridViewCheckBoxColumn
                {
                    Name = "isChecked",
                    DataPropertyName = "isChecked",
                    HeaderText = "V",
                    TrueValue = true,
                    FalseValue = false,
                    ThreeState = false
                },
                new DataGridViewTextBoxColumn
                {
                    Name = "fileName",
                    DataPropertyName = "fileName",
                    HeaderText = "שם קובץ"
                },
                new DataGridViewTextBoxColumn
                {
                    Name = "fileSize",
                    DataPropertyName = "fileSize",
                    HeaderText = "גודל"
                }
            });

            dgvAttachments.ReadOnly = true;

            dgvAttachments.BackgroundColor = SystemColors.Window;
            dgvAttachments.ForeColor = SystemColors.WindowText;
            dgvAttachments.DefaultCellStyle.BackColor = SystemColors.Window;
            dgvAttachments.DefaultCellStyle.ForeColor = SystemColors.WindowText;
            dgvAttachments.DefaultCellStyle.SelectionBackColor = SystemColors.Highlight;
            dgvAttachments.DefaultCellStyle.SelectionForeColor = SystemColors.HighlightText;
            dgvAttachments.AlternatingRowsDefaultCellStyle.BackColor = SystemColors.Control;
            dgvAttachments.AlternatingRowsDefaultCellStyle.ForeColor = SystemColors.ControlText;
            dgvAttachments.ColumnHeadersDefaultCellStyle.BackColor = SystemColors.Control;
            dgvAttachments.ColumnHeadersDefaultCellStyle.ForeColor = SystemColors.ControlText;
            dgvAttachments.EnableHeadersVisualStyles = false;
            dgvAttachments.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;

            dgvAttachments.AllowUserToAddRows = false;
        }

        private void BtnSettings_Click(object sender, EventArgs e)
        {
            using (var frm = new FormSettings(this))
            {
                frm.ShowDialog();
            }
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

        private void CmbSaveLocation_PathConfirmed(object sender, PathConfirmedEventArgs e)
        {
            if (!string.IsNullOrWhiteSpace(e.Path))
            {
                txtFullPath.Text = PathBreadcrumbHelper.FormatPathAsBreadcrumb(e.Path);
            }
        }
    }
}