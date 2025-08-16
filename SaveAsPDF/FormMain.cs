// Ignore Spelling: frm
using Microsoft.Office.Interop.Outlook;
using SaveAsPDF.Helpers;
using SaveAsPDF.Models;
using SaveAsPDF.Properties;
using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.ComponentModel; // For BindingList
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Exception = System.Exception;

namespace SaveAsPDF
{
    /// <summary>
    /// The main form for the SaveAsPDF application. Handles project, employee, and settings management,
    /// as well as user interface logic for working with Outlook mail items and attachments.
    /// </summary>
    public partial class FormMain : Form, IEmployeeRequester, INewProjectRequester, ISettingsRequester
    {
        /// <summary>
        /// The list of employees associated with the current project.
        /// </summary>
        private List<EmployeeModel> _employeesModel = new List<EmployeeModel>();

        /// <summary>
        /// The current project model.
        /// </summary>
        private ProjectModel _projectModel = new ProjectModel();

        /// <summary>
        /// The global settings model for the application.
        /// </summary>
        public static SettingsModel settingsModel = new SettingsModel();

        /// <summary>
        /// Indicates whether the data has been loaded for the current session.
        /// </summary>
        private bool _dataLoaded = false;

        /// <summary>
        /// The currently selected node in the folder tree view.
        /// </summary>
        public static TreeNode mySelectedNode;

        /// <summary>
        /// The static reference to the current Outlook mail item.
        /// </summary>
        private static MailItem _mi = null;

        /// <summary>
        /// The current Outlook mail item being processed.
        /// </summary>
        private MailItem _mailItem = ThisAddIn.TypeOfMailitem(_mi);

        /// <summary>
        /// Indicates whether the last mouse click was a double-click.
        /// </summary>
        private bool _isDoubleClick = false;

        /// <summary>
        /// The search history for project IDs.
        /// </summary>
        private List<string> searchHistory = new List<string>();

        /// <summary>
        /// The list of attachments for the current mail item.
        /// </summary>
        List<AttachmentsModel> attachmentsModels = new List<AttachmentsModel>();

        /// <summary>
        /// The settings model object for use in static contexts.
        /// </summary>
        public static object SettingModel { get; internal set; }

        /// <summary>
        /// The binding list of employees for data binding.
        /// </summary>
        private BindingList<EmployeeModel> _employeesBindingList = new BindingList<EmployeeModel>();

        // Status strip hover helpers
        private void Control_MouseEnterStatus(object sender, EventArgs e)
        {
            var ctl = sender as Control;
            if (ctl != null)
                tsslStatus.Text = ctl.Tag as string ?? string.Empty;
        }

        private void Control_MouseLeaveStatus(object sender, EventArgs e)
        {
            tsslStatus.Text = string.Empty;
        }

        private void WireStatusHelp()
        {
            // Buttons
            btnOK.Tag = "שמור ל-PDF";
            btnOK.MouseEnter += Control_MouseEnterStatus; btnOK.MouseLeave += Control_MouseLeaveStatus;

            btnCancel.Tag = "בטל וסגור";
            btnCancel.MouseEnter += Control_MouseEnterStatus; btnCancel.MouseLeave += Control_MouseLeaveStatus;

            btnSettings.Tag = "הגדרות";
            btnSettings.MouseEnter += Control_MouseEnterStatus; btnSettings.MouseLeave += Control_MouseLeaveStatus;

            btnNewProject.Tag = "פרויקט חדש";
            btnNewProject.MouseEnter += Control_MouseEnterStatus; btnNewProject.MouseLeave += Control_MouseLeaveStatus;

            btnFolders.Tag = "בחר תיקיית שורש";
            btnFolders.MouseEnter += Control_MouseEnterStatus; btnFolders.MouseLeave += Control_MouseLeaveStatus;

            btnCopyNotesToMail.Tag = "העתק הערות לפרויקט אל המייל";
            btnCopyNotesToMail.MouseEnter += Control_MouseEnterStatus; btnCopyNotesToMail.MouseLeave += Control_MouseLeaveStatus;

            btnCopyNotesToProject.Tag = "העתק הערות מהמייל אל הפרויקט";
            btnCopyNotesToProject.MouseEnter += Control_MouseEnterStatus; btnCopyNotesToProject.MouseLeave += Control_MouseLeaveStatus;

            btnStyle.Tag = "בחר גופן להערות";
            btnStyle.MouseEnter += Control_MouseEnterStatus; btnStyle.MouseLeave += Control_MouseLeaveStatus;

            RemoveEmployee.Tag = "הסר עובד מהרשימה";
            RemoveEmployee.MouseEnter += Control_MouseEnterStatus; RemoveEmployee.MouseLeave += Control_MouseLeaveStatus;

            btnPhoneBook.Tag = "בחר עובד מספר טלפונים";
            btnPhoneBook.MouseEnter += Control_MouseEnterStatus; btnPhoneBook.MouseLeave += Control_MouseLeaveStatus;

            // Inputs
            txtProjectID.Tag = "הכנס מספר פרויקט";
            txtProjectID.MouseEnter += Control_MouseEnterStatus; txtProjectID.MouseLeave += Control_MouseLeaveStatus;

            txtProjectName.Tag = "שם הפרויקט";
            txtProjectName.MouseEnter += Control_MouseEnterStatus; txtProjectName.MouseLeave += Control_MouseLeaveStatus;

            txtSubject.Tag = "נושא ההודעה";
            txtSubject.MouseEnter += Control_MouseEnterStatus; txtSubject.MouseLeave += Control_MouseLeaveStatus;

            txtFullPath.Tag = "נתיב מלא";
            txtFullPath.MouseEnter += Control_MouseEnterStatus; txtFullPath.MouseLeave += Control_MouseLeaveStatus;

            cmbSaveLocation.Tag = "בחר מיקום שמירה";
            cmbSaveLocation.MouseEnter += Control_MouseEnterStatus; cmbSaveLocation.MouseLeave += Control_MouseLeaveStatus;

            rtxtNotes.Tag = "הערות למייל";
            rtxtNotes.MouseEnter += Control_MouseEnterStatus; rtxtNotes.MouseLeave += Control_MouseLeaveStatus;

            rtxtProjectNotes.Tag = "הערות לפרויקט";
            rtxtProjectNotes.MouseEnter += Control_MouseEnterStatus; rtxtProjectNotes.MouseLeave += Control_MouseLeaveStatus;

            // Checkboxes
            chkbSendNote.Tag = "שלח הערה לראש הפרויקט";
            chkbSendNote.MouseEnter += Control_MouseEnterStatus; chkbSendNote.MouseLeave += Control_MouseLeaveStatus;

            chkbSelectAllAttachments.Tag = "בחר/הסר כל הקבצים";
            chkbSelectAllAttachments.MouseEnter += Control_MouseEnterStatus; chkbSelectAllAttachments.MouseLeave += Control_MouseLeaveStatus;

            chbOpenPDF.Tag = "פתח PDF לאחר שמירה";
            chbOpenPDF.MouseEnter += Control_MouseEnterStatus; chbOpenPDF.MouseLeave += Control_MouseLeaveStatus;

            // Lists / Trees / Grids
            tvFolders.Tag = "עץ תיקיות פרויקט";
            tvFolders.MouseEnter += Control_MouseEnterStatus; tvFolders.MouseLeave += Control_MouseLeaveStatus;

            dgvAttachments.Tag = "קבצים מצורפים";
            dgvAttachments.MouseEnter += Control_MouseEnterStatus; dgvAttachments.MouseLeave += Control_MouseLeaveStatus;

            dgvEmployees.Tag = "עובדי פרויקט";
            dgvEmployees.MouseEnter += Control_MouseEnterStatus; dgvEmployees.MouseLeave += Control_MouseLeaveStatus;

            // Tabs (optional, hover over tab control area)
            tabNotes.Tag = "כרטיסיות הערות";
            tabNotes.MouseEnter += Control_MouseEnterStatus; tabNotes.MouseLeave += Control_MouseLeaveStatus;

            tabFilesFolders.Tag = "קבצים ותיקיות";
            tabFilesFolders.MouseEnter += Control_MouseEnterStatus; tabFilesFolders.MouseLeave += Control_MouseLeaveStatus;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="FormMain"/> class.
        /// </summary>
        public FormMain()
        {
            InitializeComponent();
            
            // Use event handler for efficient loading
            Load += FormMain_Load;
            FormClosing += FormMain_FormClosing;
            
            // Setup DataGridView
            ConfigureEmployeeDataGrid();
            
            // Set up key handlers
            KeyDown += FormMain_KeyDown;
        }

        /// <summary>
        /// Configure the employees DataGridView
        /// </summary>
        private void ConfigureEmployeeDataGrid()
        {
            dgvEmployees.AutoGenerateColumns = false;
            dgvEmployees.CellValueChanged += dgvEmployees_CellValueChanged;
            dgvEmployees.CurrentCellDirtyStateChanged += dgvEmployees_CurrentCellDirtyStateChanged;
            
            // Add columns programmatically
            dgvEmployees.Columns.Clear();
            
            dgvEmployees.Columns.AddRange(new DataGridViewColumn[] {
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
                },
                new DataGridViewCheckBoxColumn
                {
                    Name = "IsLeader",
                    DataPropertyName = "IsLeader",
                    HeaderText = "ראש פרויקט",
                    ReadOnly = false // Ensure editable
                }
            });
            
            // Bind the DataGridView to the BindingList
            dgvEmployees.DataSource = _employeesBindingList;

            // Make all columns read-only by default
            dgvEmployees.ReadOnly = true;

            // Allow editing only in the "IsLeader" column
            dgvEmployees.Columns["IsLeader"].ReadOnly = false;
        }

        /// <summary>
        /// Loads the subject of the current mail item, or a default value if unavailable.
        /// </summary>
        /// <returns>The subject string.</returns>
        private string LoadEmailSubject() => _mailItem?.Subject ?? "Default Email Subject";

        /// <summary>
        /// Handles the form load event. Initializes UI elements, loads settings, and populates controls.
        /// </summary>
        private void FormMain_Load(object sender, EventArgs e)
        {
            // Apply context menus to improve UX
            SetupContextMenus();
            
            // Set up attachment selection UI
            chkbSelectAllAttachments.Checked = true;
            chkbSelectAllAttachments.Text = "הסר הכל";

            // Set up project ID auto-complete
            ConfigureProjectIdAutoComplete();
            
            // Load initial data
            txtSubject.Text = LoadEmailSubject();
            LoadSearchHistory();
            settingsModel = SettingsHelpers.LoadProjectSettings();

            // Initialize folder tree view with logical drives
            PopulateDriveNodes();

            // NEW: wire status help
            WireStatusHelp();
        }
        
        /// <summary>
        /// Configures context menus for various controls
        /// </summary>
        private void SetupContextMenus()
        {
            // Context menus for rich text and text controls
            rtxtNotes.EnableContextMenu();
            rtxtProjectNotes.EnableContextMenu();
            txtFullPath.EnableContextMenu();
            txtProjectID.EnableContextMenu();
            txtProjectName.EnableContextMenu();
            tvFolders.EnableContextMenu();
        }
        
        /// <summary>
        /// Configure Project ID AutoComplete functionality
        /// </summary>
        private void ConfigureProjectIdAutoComplete()
        {
            txtProjectID.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            txtProjectID.AutoCompleteSource = AutoCompleteSource.CustomSource;
        }
        
        /// <summary>
        /// Populates the folder tree with logical drives
        /// </summary>
        private void PopulateDriveNodes()
        {
            tvFolders.Nodes.Clear();
            
            foreach (string drive in Environment.GetLogicalDrives())
            {
                try
                {
                    var di = new DriveInfo(drive);
                    int driveImage = 2;
                    
                    // Set appropriate icon based on drive type
                    if (di.DriveType == DriveType.CDRom)
                        driveImage = 3;
                    else if (di.DriveType == DriveType.Network)
                        driveImage = 6;
                    else if (di.DriveType == DriveType.NoRootDirectory || di.DriveType == DriveType.Unknown)
                        driveImage = 8;

                    var node = new TreeNode(drive.Substring(0, 1), driveImage, driveImage) { Tag = drive };
                    
                    // Add placeholder node for expandable drives
                    if (di.IsReady) 
                        node.Nodes.Add("...");
                        
                    tvFolders.Nodes.Add(node);
                }
                catch
                {
                    // Skip drives that throw exceptions
                }
            }
        }

        /// <summary>
        /// Updates the auto-complete source for the project ID textbox with the specified text.
        /// </summary>
        /// <param name="text">The text to add to the auto-complete source.</param>
        private void UpdateAutoCompleteSource(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
                return;

            // Remove if already exists to avoid duplicates and keep most recent at the top
            searchHistory.Remove(text);
            searchHistory.Insert(0, text);

            // Limit history to 10 items (or any desired max)
            int maxCount = 10;
            if (searchHistory.Count > maxCount)
                searchHistory = searchHistory.Take(maxCount).ToList();

            // Update settings
            if (Settings.Default.LastProjects == null)
                Settings.Default.LastProjects = new StringCollection();
            Settings.Default.LastProjects.Clear();
            Settings.Default.LastProjects.AddRange(searchHistory.ToArray());
            Settings.Default.Save();

            // Update the textbox's auto-complete source immediately
            txtProjectID.AutoCompleteCustomSource.Clear();
            txtProjectID.AutoCompleteCustomSource.AddRange(searchHistory.ToArray());
        }

        /// <summary>
        /// Loads the search history from application settings into the project ID auto-complete source.
        /// </summary>
        private void LoadSearchHistory()
        {
            if (Settings.Default.LastProjects == null)
                Settings.Default.LastProjects = new StringCollection();
            searchHistory = Settings.Default.LastProjects.Cast<string>().Distinct().ToList();
            txtProjectID.AutoCompleteCustomSource.AddRange(searchHistory.ToArray());
        }

        /// <summary>
        /// Handles the Cancel button click event. Closes the form.
        /// </summary>
        private void btnCancel_Click(object sender, EventArgs e) => Close();

        /// <summary>
        /// Handles the Folders button click event. Opens a folder picker and updates the folder tree view.
        /// </summary>
        private void btnFolders_Click(object sender, EventArgs e)
        {
            var dialog = new FolderPicker
            {
                InputPath = settingsModel.ProjectRootFolder.Exists ? settingsModel.ProjectRootFolder.FullName : settingsModel.RootDrive
            };
            if (dialog.ShowDialog(Handle) == true)
            {
                tvFolders.Nodes.Clear();
                tvFolders.Nodes.Add(TreeHelpers.TraverseDirectory(dialog.ResultPath, 1));
            }
        }

        /// <summary>
        /// Handles the KeyDown event for the project ID textbox. Selects the OK button on Enter.
        /// </summary>
        private void txtProjectID_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                btnOK.Select();
        }

        /// <summary>
        /// Handles the first run scenario by prompting the user to select a root drive.
        /// </summary>
        private void HandleFirstRun()
        {
            XMessageBox.Show(
                "זהו השימוש הראשון בתוכנה",
                "SaveAsPDF",
                XMessageBoxButtons.OK,
                XMessageBoxIcon.Information,
                XMessageAlignment.Right,
                XMessageLanguage.Hebrew);
            var dialog = new FolderPicker { InputPath = settingsModel.RootDrive };
            if (dialog.ShowDialog(Handle) == true)
                settingsModel.RootDrive = dialog.InputPath;
        }

        /// <summary>
        /// Processes the entered project ID, loads settings, and updates the UI.
        /// </summary>
        /// <param name="projectID">The project ID to process.</param>
        private void ProcessProjectID(string projectID)
        {
            if (!projectID.SafeProjectID())
            {
                XMessageBox.Show(
                    Resources.InvalidProjectIDMessage ?? "מספר פרויקט לא חוקי",
                    Resources.ErrorTitle ?? "שגיאה",
                    XMessageBoxButtons.OK,
                    XMessageBoxIcon.Error,
                    XMessageAlignment.Right,
                    XMessageLanguage.Hebrew);
                return;
            }
            try
            {
                settingsModel = SettingsHelpers.LoadProjectSettings(projectID);
                LoadProjectData();
                LoadEmployeeData();
                UpdateUI();
                
                // Check for duplicate project IDs in the path and fix if needed
                if (settingsModel.ProjectRootFolder != null && settingsModel.ProjectRootFolder.Exists)
                {
                    string projectPath = settingsModel.ProjectRootFolder.FullName;
                    string folderStructure = $"\\{projectID}\\{projectID}\\";

                    if (projectPath.Contains(folderStructure))
                    {
                        // Fix the path to avoid duplicate project IDs
                        projectPath = projectPath.Replace(folderStructure, $"\\{projectID}\\");
                    }
                    
                    cmbSaveLocation.Text = projectPath;
                }
                else
                {
                    cmbSaveLocation.Text = settingsModel.DefaultSavePath;
                }
                
                _dataLoaded = true;
            }
            catch (FileNotFoundException ex)
            {
                XMessageBox.Show(
                    $"קובץ נדרש לא נמצא: {ex.Message}",
                    "קובץ לא נמצא",
                    XMessageBoxButtons.OK,
                    XMessageBoxIcon.Warning,
                    XMessageAlignment.Right,
                    XMessageLanguage.Hebrew);
            }
            catch (UnauthorizedAccessException ex)
            {
                XMessageBox.Show(
                    $"הגישה לקובץ או לתיקייה נדחתה: {ex.Message}",
                    "גישה נדחתה",
                    XMessageBoxButtons.OK,
                    XMessageBoxIcon.Error,
                    XMessageAlignment.Right,
                    XMessageLanguage.Hebrew);
            }
            catch (Exception ex)
            {
                XMessageBox.Show(
                    $"אירעה שגיאה בלתי צפויה בעת עיבוד מספר הפרויקט: {ex.Message}",
                    "שגיאה",
                    XMessageBoxButtons.OK,
                    XMessageBoxIcon.Error,
                    XMessageAlignment.Right,
                    XMessageLanguage.Hebrew);
            }
        }

        /// <summary>
        /// Loads the project data from the XML file or sets default values if unavailable.
        /// </summary>
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
                        chkbSendNote.Checked = _projectModel.NoteToProjectLeader;
                        rtxtProjectNotes.Text = _projectModel.ProjectNotes;
                    }
                }
                catch (Exception ex)
                {
                    XMessageBox.Show(
                        $"שגיאה בטעינת נתוני הפרויקט: {ex.Message}\nהקובץ יוחלף בערכים ברירת מחדל.",
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

        /// <summary>
        /// Sets the project model to default values and updates the UI.
        /// </summary>
        private void SetDefaultProjectModel()
        {
            _projectModel = new ProjectModel
            {
                ProjectName = "פרויקט ברירת מחדל",
                ProjectNumber = "0000",
                NoteToProjectLeader = false,
                DefaultSaveFolder = settingsModel.DefaultSavePath,
                ProjectNotes = "הערות ברירת מחדל",
                LastSavePath = settingsModel.DefaultSavePath
            };
            settingsModel.XmlProjectFile.ProjectModelToXmlFile(_projectModel);
            txtProjectName.Text = _projectModel.ProjectName;
            chkbSendNote.Checked = _projectModel.NoteToProjectLeader;
            rtxtProjectNotes.Text = _projectModel.ProjectNotes;
        }

        /// <summary>
        /// Loads the employee data from the XML file and populates the data grid view.
        /// </summary>
        private void LoadEmployeeData()
        {
            _employeesBindingList.Clear();
            if (File.Exists(settingsModel.XmlProjectFile))
            {
                var loaded = settingsModel.XmlEmployeesFile.XmlEmployeesFileToModel();
                if (loaded != null)
                {
                    foreach (var em in loaded)
                        _employeesBindingList.Add(em);
                }
            }
        }

        /// <summary>
        /// Updates the UI elements based on the current settings and project data.
        /// </summary>
        private void UpdateUI()
        {
            try
            {
                // Clear and populate the save location combo box
                cmbSaveLocation.Items.Clear();
                
                // First add the project root folder path directly
                if (settingsModel.ProjectRootFolder != null && settingsModel.ProjectRootFolder.Exists)
                {
                    // Check for duplicate project IDs in the path
                    string projectID = txtProjectID.Text;
                    string projectPath = settingsModel.ProjectRootFolder.FullName;
                    
                    // Make sure projectID isn't duplicated (e.g. "C:\10\1000\1000\")
                    string folderStructure = $"\\{projectID}\\{projectID}\\";
                    if (projectPath.Contains(folderStructure))
                    {
                        projectPath = projectPath.Replace(folderStructure, $"\\{projectID}\\");
                    }
                    
                    cmbSaveLocation.Items.Add(projectPath);
                }
                
                // Now load any additional paths from the tree file
                if (File.Exists(settingsModel.DefaultTreeFile))
                {
                    cmbSaveLocation.LoadComboBoxWithPaths(settingsModel.DefaultTreeFile, txtProjectID.Text);
                }
                
                cmbSaveLocation.CustomizeComboBox();
                
                if (cmbSaveLocation.Items.Count > 0)
                {
                    // Try to select the default save path
                    if (!string.IsNullOrEmpty(settingsModel.DefaultSavePath))
                    {
                        // Try to find and select the default save path
                        for (int i = 0; i < cmbSaveLocation.Items.Count; i++)
                        {
                            if (string.Equals(cmbSaveLocation.Items[i].ToString(), settingsModel.DefaultSavePath, StringComparison.OrdinalIgnoreCase))
                            {
                                cmbSaveLocation.SelectedIndex = i;
                                break;
                            }
                        }
                        
                        // If we couldn't find it, set it as the text anyway
                        if (cmbSaveLocation.SelectedIndex < 0)
                        {
                            cmbSaveLocation.Text = settingsModel.DefaultSavePath;
                        }
                    }
                    else
                    {
                        // Default to first item (which should be the project root folder)
                        cmbSaveLocation.SelectedIndex = 0;
                    }
                }

                // Update the tree view
                tvFolders.Nodes.Clear();
                if (settingsModel.ProjectRootFolder.Exists)
                {
                    tvFolders.Nodes.Add(TreeHelpers.CreateDirectoryNode(settingsModel.ProjectRootFolder));
                    tvFolders.ExpandAll();
                    tvFolders.SelectedNode = tvFolders.Nodes[0];
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
                txtFullPath.Text = settingsModel.ProjectRootFolder.FullName;
                btnOK.Focus();
            }
            catch (Exception ex)
            {
                XMessageBox.Show(
                    $"אירעה שגיאה בעת עדכון הממשק: {ex.Message}",
                    "שגיאה",
                    XMessageBoxButtons.OK,
                    XMessageBoxIcon.Error,
                    XMessageAlignment.Right,
                    XMessageLanguage.Hebrew);
            }
        }

        /// <summary>
        /// Processes the given mail item and loads its attachments into the UI.
        /// </summary>
        /// <param name="mailItem">The Outlook mail item to process.</param>
        private void ProcessMailItem(MailItem mailItem)
        {
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
            dgvAttachments.DataSource = attachmentsModels;
            dgvAttachments.Columns[0].Visible = false;
            dgvAttachments.Columns[1].HeaderText = "V";
            dgvAttachments.Columns[1].ReadOnly = false;
            dgvAttachments.Columns[2].HeaderText = "שם קובץ";
            dgvAttachments.Columns[2].ReadOnly = true;
            dgvAttachments.Columns[3].HeaderText = "גודל";
            dgvAttachments.Columns[3].ReadOnly = true;
        }

        /// <summary>
        /// Shows an error message if the selected item is not a valid mail item.
        /// </summary>
        private void ShowInvalidMailItemError()
        {
            XMessageBox.Show(
                "יש לבחור הודעות דואר אלקטרוני בלבד",
                "SaveAsPDF",
                XMessageBoxButtons.OK,
                XMessageBoxIcon.Error,
                XMessageAlignment.Right,
                XMessageLanguage.Hebrew);
            Close();
        }

        /// <summary>
        /// Clears the form fields and resets the UI for a new operation.
        /// </summary>
        private void ClearForm()
        {
            txtProjectName.Clear();
            txtFullPath.Clear();
            cmbSaveLocation.Items.Clear();
            rtxtNotes.Clear();
            rtxtProjectNotes.Clear();
            dgvAttachments.DataSource = null;
            dgvEmployees.DataSource = null;
            tvFolders.DataBindings.Clear();
        }

        /// <summary>
        /// Handles the OK button click event. Validates the save path, generates HTML, saves the mail as PDF, and optionally opens the PDF.
        /// </summary>
        private void btnOK_Click(object sender, EventArgs e)
        {
            string sPath = cmbSaveLocation.Text;
            if (string.IsNullOrEmpty(sPath))
            {
                var dialog = new FolderPicker { InputPath = settingsModel.RootDrive };
                if (dialog.ShowDialog(Handle) == true)
                    sPath = dialog.ResultPath;
            }
            if (!string.IsNullOrEmpty(sPath))
            {
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

            string htmlContent = HtmlHelper.GenerateHtmlContent(
                sPath,
                _employeesModel,
                attachmentsModels,
                txtProjectName.Text,
                txtProjectID.Text,
                rtxtNotes.Text,
                Environment.UserName
            );

            _mailItem.HTMLBody = htmlContent + _mailItem.HTMLBody;
            _mailItem.SaveToPDF(sPath);
            _mailItem.Save();
            Close();

            if (chbOpenPDF.Checked)
            {
                string pdfFilePath = Path.Combine(sPath, $"{_mailItem.Subject}.pdf");
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

        /// <summary>
        /// Opens the settings form as a modal dialog.
        /// </summary>
        private void BtnSettings_Click(object sender, EventArgs e)
        {
            using (var frm = new FormSettings(this))
            {
                frm.ShowDialog();
            }
        }

        /// <summary>
        /// Handles the double-click event on the attachments data grid view. Shows file details.
        /// </summary>
        private void dgvAttachments_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            XMessageBox.Show(
                dgvAttachments.CurrentCell.Value.ToString(),
                "פרטי קובץ",
                XMessageBoxButtons.OK,
                XMessageBoxIcon.Information,
                XMessageAlignment.Right,
                XMessageLanguage.Hebrew);
        }

        /// <summary>
        /// Handles the style button click event. Opens a font dialog for the notes rich text box.
        /// </summary>
        private void btnStyle_Click(object sender, EventArgs e)
        {
            if (dlgFont.ShowDialog() == DialogResult.OK)
                rtxtNotes.SelectionFont = dlgFont.Font;
        }

        /// <summary>
        /// Opens the contacts form as a modal dialog.
        /// </summary>
        private void btnPhoneBook_Click(object sender, EventArgs e)
        {
            using (var frmContacts = new FormContacts(this))
            {
                frmContacts.ShowDialog(this);
            }
        }

        /// <summary>
        /// Handles the send note checkbox change event. Sends an email to each employee.
        /// </summary>
        private void chkbSendNote_CheckedChanged(object sender, EventArgs e)
        {
            foreach (var employee in _employeesModel)
                SendEmailToEmployee(employee.EmailAddress);
        }

        /// <summary>
        /// Sends an email notification to the specified employee.
        /// </summary>
        /// <param name="EmailAddress">The email address of the employee.</param>
        private void SendEmailToEmployee(string EmailAddress) =>
            XMessageBox.Show(
                $"Send email to {EmailAddress}",
                "SaveAsPDF",
                XMessageBoxButtons.OK,
                XMessageBoxIcon.Information,
                XMessageAlignment.Right,
                XMessageLanguage.Hebrew);

        /// <summary>
        /// Handles the select all attachments checkbox change event. Updates the selection state of all attachments.
        /// </summary>
        private void chkbSelectAllAttachments_CheckedChanged(object sender, EventArgs e)
        {
            chkbSelectAllAttachments.Text = chkbSelectAllAttachments.Checked ? "הסר הכל" : "בחר הכל";
            if (dgvAttachments.RowCount != 0)
            {
                foreach (DataGridViewRow row in dgvAttachments.Rows)
                    row.Cells[1].Value = chkbSelectAllAttachments.Checked;
            }
        }

        /// <summary>
        /// Called when the settings form completes. Updates the settings model.
        /// </summary>
        /// <param name="settings">The updated settings model.</param>
        public void SettingsComplete(SettingsModel settings)
        {
            settingsModel = SettingsHelpers.LoadSettingsToModel(settings);
        }

        /// <summary>
        /// Called when the contacts form completes. Adds a new employee if not already present.
        /// </summary>
        /// <param name="model">The employee model to add.</param>
        public void EmployeeComplete(EmployeeModel model)
        {
            if (!_employeesBindingList.Any(e => e.EmailAddress == model.EmailAddress))
            {
                _employeesBindingList.Add(model);
                SaveEmployeesToXml();
            }
        }

        // This event ensures checkbox changes are committed immediately
        private void dgvEmployees_CurrentCellDirtyStateChanged(object sender, EventArgs e)
        {
            if (dgvEmployees.IsCurrentCellDirty && dgvEmployees.CurrentCell is DataGridViewCheckBoxCell)
            {
                dgvEmployees.CommitEdit(DataGridViewDataErrorContexts.Commit);
            }
        }

        // This event updates the EmployeeModel and XML when IsLeader is changed
        private void dgvEmployees_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0 && dgvEmployees.Columns[e.ColumnIndex].Name == "IsLeader")
            {
                SaveEmployeesToXml();
            }
        }

        // Fix for CS1501: Adjusting the method call to match the correct signature based on the provided type signatures.
        private void SaveEmployeesToXml()
        {
            if (!string.IsNullOrEmpty(settingsModel.XmlEmployeesFile))
            {
                XmlFileHelper.EmployeesModelToXmlFile(settingsModel.XmlEmployeesFile, _employeesModel);
            }
        }

        /// <summary>
        /// Handles the BeforeExpand event for the folder tree view. Cancels expansion on double-click.
        /// </summary>
        private void tvFolders_BeforeExpand(object sender, TreeViewCancelEventArgs e)
        {
            if (_isDoubleClick && e.Action == TreeViewAction.Expand)
                e.Cancel = true;
        }

        /// <summary>
        /// Handles the remove employee button click event. Removes selected employees from the list.
        /// </summary>
        private void RemoveEmployee_Click(object sender, EventArgs e)
        {
            int selectedRowCount = dgvEmployees.Rows.GetRowCount(DataGridViewElementStates.Selected);
            if (selectedRowCount > 0)
            {
                for (int i = 0; i < selectedRowCount; i++)
                    dgvEmployees.Rows.RemoveAt(dgvEmployees.SelectedRows[0].Index);
            }
            else
            {
                XMessageBox.Show(
                    "לא נבחרו עובדים למחיקה.",
                    "שגיאה",
                    XMessageBoxButtons.OK,
                    XMessageBoxIcon.Warning,
                    XMessageAlignment.Right,
                    XMessageLanguage.Hebrew);
            }
        }

        /// <summary>
        /// Opens the new project form as a modal dialog.
        /// </summary>
        private void btnNewProject_Click(object sender, EventArgs e)
        {
            using (var frm = new FormNewProject(this))
            {
                frm.ShowDialog(this);
            }
        }

        /// <summary>
        /// Called when the new project form completes. Updates the project model and UI.
        /// </summary>
        /// <param name="model">The new project model.</param>
        public void NewProjectComplete(ProjectModel model)
        {
            _projectModel = model;
            txtProjectID.Text = _projectModel.ProjectNumber;
            txtProjectName.Text = _projectModel.ProjectName;
            rtxtProjectNotes.Text = _projectModel.ProjectNotes;
        }

        /// <summary>
        /// Handles the mouse hover event for the project ID textbox. Updates the status label.
        /// </summary>
        private void txtProjectID_MouseHover(object sender, EventArgs e)
        {
            tsslStatus.Enabled = true;
            tsslStatus.Text = "מספר פרויקט כפי שמופיע במסטרפלן";
        }

        /// <summary>
        /// Copies the project notes to the mail notes rich text box.
        /// </summary>
        private void btnCopyNotesToMail_Click(object sender, EventArgs e)
        {
            rtxtNotes.Text += $"\n {rtxtProjectNotes.Text}";
        }

        /// <summary>
        /// Copies the mail notes to the project notes rich text box.
        /// </summary>
        private void btnCopyNotesToProject_Click(object sender, EventArgs e)
        {
            rtxtProjectNotes.Text += $"\n {rtxtNotes.Text}";
        }

        /// <summary>
        /// Handles the AfterSelect event for the folder tree view. Updates the selected node.
        /// </summary>
        private void tvFolders_AfterSelect(object sender, TreeViewEventArgs e)
        {
            mySelectedNode = e.Node;
        }

        /// <summary>
        /// Handles the AfterLabelEdit event for the folder tree view. Validates and renames the node.
        /// </summary>
        private void tvFolders_AfterLabelEdit(object sender, NodeLabelEditEventArgs e)
        {
            if (e.Label == null || string.IsNullOrWhiteSpace(e.Label.SafeFolderName()))
            {
                e.CancelEdit = true;
                XMessageBox.Show(
                    "שם לא חוקי.\n לא ניתן ליצור שם ריק. חובה תו אחד לפחות",
                    "עריכת שם",
                    XMessageBoxButtons.OK,
                    XMessageBoxIcon.Error,
                    XMessageAlignment.Right,
                    XMessageLanguage.Hebrew);
                return;
            }
            string nodeNewLabel = e.Label.SafeFolderName();
            if (nodeNewLabel.IndexOfAny(new char[] { '\\', '/', ':', '*', '?', '<', '>', '|', '"' }) != -1)
            {
                e.CancelEdit = true;
                XMessageBox.Show(
                    "שם לא חוקי.\nאין להשתמש בתווים הבאים \n'\\', '/', ':', '*', '?', '<', '>', '|' '\"' ",
                    "עריכת שם",
                    XMessageBoxButtons.OK,
                    XMessageBoxIcon.Error,
                    XMessageAlignment.Right,
                    XMessageLanguage.Hebrew);
                return;
            }
            try
            {
                string oldPath = Path.Combine(settingsModel.ProjectRootFolder.Parent.FullName, e.Node.FullPath);
                string newPath = Path.Combine(settingsModel.ProjectRootFolder.Parent.FullName, e.Node.Parent.FullPath, nodeNewLabel);
                DirectoryInfo directoryInfo = new DirectoryInfo(oldPath);
                directoryInfo.RenameDirectory(newPath);
                e.Node.Text = nodeNewLabel;
            }
            catch (Exception ex)
            {
                e.CancelEdit = true;
                XMessageBox.Show(
                    $"שגיאה בשינוי שם התיקייה: {ex.Message}\n{Path.Combine(settingsModel.ProjectRootFolder.Parent.FullName, e.Node.FullPath)}",
                    "SaveAsPDF:tvFolders_AfterLabelEdit",
                    XMessageBoxButtons.OK,
                    XMessageBoxIcon.Error,
                    XMessageAlignment.Right,
                    XMessageLanguage.Hebrew);
            }
        }

        /// <summary>
        /// Handles the Open PDF checkbox change event. Updates the settings model.
        /// </summary>
        private void chbOpenPDF_CheckedChanged(object sender, EventArgs e)
        {
            settingsModel.OpenPDF = chbOpenPDF.Checked;
        }

        /// <summary>
        /// Handles the double-click event on the folder tree view. Opens the folder in Explorer and updates the save location.
        /// </summary>
        private void tvFolders_NodeMouseDoubleClick(object sender, TreeNodeMouseClickEventArgs e)
        {
            string path = $@"{settingsModel.ProjectRootFolder.Parent.FullName}\{e.Node.FullPath}";
            System.Diagnostics.Process.Start("explorer.exe", path);
            cmbSaveLocation.SelectedText = path;
        }

        /// <summary>
        /// Handles the mouse click event on the folder tree view. Updates the save location.
        /// </summary>
        private void tvFolders_NodeMouseClick(object sender, TreeNodeMouseClickEventArgs e)
        {
            string path = $@"{settingsModel.ProjectRootFolder.Parent.FullName}\{e.Node.FullPath}";
            cmbSaveLocation.Select();
            cmbSaveLocation.SelectedItem = null;
            cmbSaveLocation.SelectedText = path;
        }

        /// <summary>
        /// Handles the Validating event for the project ID textbox. Validates the project ID and updates the UI.
        /// </summary>
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

        /// <summary>
        /// Handles the Validated event for the project ID textbox. Processes the project ID and loads data.
        /// </summary>
        private void txtProjectID_Validated(object sender, EventArgs e)
        {
            errorProviderMain.SetError(txtProjectID, string.Empty);
            tsslStatus.Text = errorProviderMain.GetError(txtProjectID);

            string projectID = txtProjectID.Text;
            if (!string.IsNullOrWhiteSpace(projectID))
            {
                // Save the valid project ID to auto-complete history
                UpdateAutoCompleteSource(projectID);
            }

            // Check if the project folder exists before proceeding
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
                return;
            }

            ProcessProjectID(projectID);
            if (string.IsNullOrEmpty(settingsModel.RootDrive))
                HandleFirstRun();
            
            // Ensure the project root path is displayed in the ComboBox without duplicate project IDs
            if (settingsModel.ProjectRootFolder != null && settingsModel.ProjectRootFolder.Exists)
            {
                string projectPath = settingsModel.ProjectRootFolder.FullName;
                string folderStructure = $"\\{projectID}\\{projectID}\\";

                if (projectPath.Contains(folderStructure))
                {
                    // Fix the path to avoid duplicate project IDs
                    projectPath = projectPath.Replace(folderStructure, $"\\{projectID}\\");
                }

                cmbSaveLocation.Text = projectPath;
            }
            
            if (_mailItem is MailItem mailItem)
                ProcessMailItem(_mailItem);
            else
                ShowInvalidMailItemError();
        }

        /// <summary>
        /// Handles the FormClosing event. Prompts the user for confirmation before exiting.
        /// </summary>
        private void FormMain_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (e.CloseReason == CloseReason.UserClosing)
            {
                if (XMessageBox.Show(
                        "האם לצאת מהיישום?",
                        "SaveAsPDF",
                        XMessageBoxButtons.YesNo,
                        XMessageBoxIcon.Question,
                        XMessageAlignment.Right,
                        XMessageLanguage.Hebrew) == DialogResult.No)
                {
                    e.Cancel = true;
                }
            }
        }

        /// <summary>
        /// Handles the BeforeCollapse event for the folder tree view. Cancels collapse on double-click.
        /// </summary>
        private void tvFolders_BeforeCollapse(object sender, TreeViewCancelEventArgs e)
        {
            if (_isDoubleClick && e.Action == TreeViewAction.Collapse)
                e.Cancel = true;
        }

        /// <summary>
        /// Handles the MouseDown event for the folder tree view. Tracks double-click state.
        /// </summary>
        private void tvFolders_MouseDown(object sender, MouseEventArgs e)
        {
            _isDoubleClick = e.Clicks > 1;
        }

        /// <summary>
        /// Handles the KeyDown event for the form. Provides keyboard shortcuts for folder operations.
        /// </summary>
        private void FormMain_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.F2 && tvFolders.SelectedNode != null)
            {
                tvFolders.SelectedNode.BeginEdit();
            }
            else if (e.KeyCode == Keys.Delete && tvFolders.SelectedNode != null && tvFolders.SelectedNode.Nodes.Count == 0)
            {
                try
                {
                    DirectoryInfo directoryInfo = new DirectoryInfo(Path.Combine(settingsModel.ProjectRootFolder.Parent.FullName, tvFolders.SelectedNode.FullPath));
                    directoryInfo.Delete(true);
                    tvFolders.SelectedNode.Remove();
                }
                catch (Exception ex)
                {
                    XMessageBox.Show(
                        $"אירעה שגיאה בעת מחיקת התיקייה: {ex.Message}",
                        "שגיאה",
                        XMessageBoxButtons.OK,
                        XMessageBoxIcon.Error,
                        XMessageAlignment.Right,
                        XMessageLanguage.Hebrew);
                }
            }
            else if (e.KeyCode == Keys.F5)
            {
                tvFolders.Nodes.Clear();
                tvFolders.Nodes.Add(TreeHelpers.TraverseDirectory(settingsModel.ProjectRootFolder.FullName, 1));
            }
            else if (e.KeyCode == Keys.Escape)
            {
                if (txtProjectID.Text.Length == 0)
                    Close();
                txtProjectID.Clear();
            }
        }

        /// <summary>
        /// Handles the SelectedValueChanged event for the save location combo box. Updates the project model.
        /// </summary>
        private void cmbSaveLocation_SelectedValueChanged(object sender, EventArgs e)
        {
            if (_projectModel == null)
            {
                _projectModel = new ProjectModel
                {
                    ProjectName = "Default Project",
                    ProjectNumber = "0000",
                    NoteToProjectLeader = false,
                    DefaultSaveFolder = settingsModel.DefaultSavePath,
                    ProjectNotes = "Default notes",
                    LastSavePath = settingsModel.DefaultSavePath
                };
            }
            _projectModel.LastSavePath = cmbSaveLocation.SelectedItem?.ToString();
        }

        /// <summary>
        /// Handles the TextUpdate event for the save location combo box. Updates the breadcrumb path.
        /// </summary>
        private void cmbSaveLocation_TextUpdate(object sender, EventArgs e)
        {
            cmbSaveLocation.SetBreadcrumbPath();
        }
    }
}