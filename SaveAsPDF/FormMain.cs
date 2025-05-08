// Ignore Spelling: frm
using Microsoft.Office.Interop.Outlook;
using SaveAsPDF.Helpers;

using SaveAsPDF.Models;
using SaveAsPDF.Properties;
using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;

using Exception = System.Exception;

namespace SaveAsPDF
{
    public partial class FormMain : Form, IEmployeeRequester, INewProjectRequester, ISettingsRequester
    {
        //Preparing the models - Private fields 
        private List<EmployeeModel> _employeesModel = new List<EmployeeModel>();
        private ProjectModel _projectModel = new ProjectModel();

        public static SettingsModel settingsModel = new SettingsModel(); //this is a property 


        // construct the full path for everything
        //private DirectoryInfo _xmlSaveAsPdfFolder;
        //private string _xmlProjectFile;
        //private string _xmlEmploeeysFile;

        private bool _dataLoaded = false;

        public static TreeNode mySelectedNode;

        private static MailItem _mi = null;
        private MailItem _mailItem = ThisAddIn.TypeOfMailitem(_mi);

        private bool _isDoubleClick = false;

        //AutoComplete search history for txtProjectID
        private List<string> searchHistory = new List<string>();


        //DocumentModel oDoc = new DocumentModel();

        //List<Attachment> attachments = new List<Attachment>();
        List<AttachmentsModel> attachmentsModels = new List<AttachmentsModel>();
        /// <summary>
        /// The model for the settings form.
        /// </summary>
        public static object SettingModel { get; internal set; }


        public FormMain()
        {
            InitializeComponent();
            this.Load += new EventHandler(FormMain_Load);
        }
        /// <summary>
        /// Load the email subject from the mail item.
        /// </summary>
        /// <returns> mailItem.subject </returns>
        private string LoadEmailSubject()
        {
            // load the email subject from mail item
            if (_mailItem != null)
                return _mailItem.Subject;
            else
                return "Default Email Subject";

        }
        /// <summary>
        /// Handles the form load event.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void FormMain_Load(object sender, EventArgs e)
        {

            //Employees dataGridView columns headers 
            dgvEmployees.Columns[0].Visible = false;
            dgvEmployees.Columns[1].HeaderText = "שם פרטי";
            dgvEmployees.Columns[2].HeaderText = "שם משפחה";
            dgvEmployees.Columns[3].HeaderText = "אימייל";
            //dgvEmployees.Columns[4].Visible = false;

            //Load the context menu to the rich text-boxes 
            rtxtNotes.EnableContextMenu();
            rtxtProjectNotes.EnableContextMenu();
            txtFullPath.EnableContextMenu();
            txtProjectID.EnableContextMenu();
            txtProjectName.EnableContextMenu();
            tvFolders.EnableContextMenu();

            chkbSelectAllAttachments.Checked = true;
            chkbSelectAllAttachments.Text = "הסר הכל";


            txtProjectID.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            txtProjectID.AutoCompleteSource = AutoCompleteSource.CustomSource;

            // Load the email subject into txtSubject
            txtSubject.Text = LoadEmailSubject();

            LoadSearchHistory();

            settingsModel = SettingsHelpers.LoadProjectSettings();

            //get a list of the drives for the folders tree 
            string[] drives = Environment.GetLogicalDrives();
            foreach (string drive in drives)
            {
                DriveInfo di = new DriveInfo(drive);
                int driveImage;
                switch (di.DriveType)    //set the drive's icon
                {
                    case DriveType.CDRom:
                        driveImage = 3;
                        break;
                    case DriveType.Network:
                        driveImage = 6;
                        break;
                    case DriveType.NoRootDirectory:
                        driveImage = 8;
                        break;
                    case DriveType.Unknown:
                        driveImage = 8;
                        break;
                    default:
                        driveImage = 2;
                        break;
                }

                TreeNode node = new TreeNode(drive.Substring(0, 1), driveImage, driveImage);
                node.Tag = drive;

                if (di.IsReady == true)
                    node.Nodes.Add("...");

                tvFolders.Nodes.Add(node);

                KeyDown += FormMain_KeyDown; //add the key-down event to the form
            }

        }
        /// <summary>
        /// Updates the auto complete source with the current text in the project ID textbox.
        /// <param name="text"> The text to be added to the auto-complete source.</param>
        /// </summary>
        private void UpdateAutoCompleteSource(string text)
        {
            if (!searchHistory.Contains(text)) //not adding duplicates
            {
                searchHistory.Add(text);
                if (Settings.Default.LastProjects == null)
                {
                    Settings.Default.LastProjects = new StringCollection();
                }
                Settings.Default.LastProjects.AddRange(searchHistory.ToArray());
                Settings.Default.Save();
            }
        }
        /// <summary>
        /// Loads the search history from settings and adds it to the auto-complete source of the project ID textbox.
        /// </summary>
        private void LoadSearchHistory()
        {
            if (Settings.Default.LastProjects == null)
            {
                Settings.Default.LastProjects = new StringCollection();
            }
            //remove duplicates fromSettings.Default.LastProjects
            searchHistory = Settings.Default.LastProjects.Cast<string>().Distinct().ToList();
            txtProjectID.AutoCompleteCustomSource.AddRange(searchHistory.ToArray());
        }
        /// <summary>
        /// Handles the key down event for the form.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnCancel_Click(object sender, EventArgs e)
        {
            //close the form - do  nothing
            Close();
        }
        /// <summary>
        /// Handles the key down event for the form.
        /// <param name="sender"></param>"
        /// </summary>
        /// <param name="sender"> The source of the event.</param>
        /// <param name="e"> The event data.</param>
        private void btnFolders_Click(object sender, EventArgs e)
        {
            var Dialog = new FolderPicker();

            if (!settingsModel.ProjectRootFolder.Exists)
            {
                Dialog.InputPath = settingsModel.RootDrive;
            }
            else
            {
                Dialog.InputPath = settingsModel.ProjectRootFolder.FullName;
            }

            if (Dialog.ShowDialog(Handle) == true)
            {
                tvFolders.Nodes.Clear();
                tvFolders.Nodes.Add(TreeHelpers.TraverseDirectory(Dialog.ResultPath, 1));
            }
        }
        private void txtProjectID_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                //when the user press Enter key, we will process the project ID (trigger the Validating event)
                btnOK.Select();
            }
        }


        /// <summary>
        /// If rootDrive (j:\) does not exist on settings.settings or folder not exist, we shout check it 
        /// </summary>
        private void HandleFirstRun()
        {
            MessageBox.Show("This is the first run");
            var dialog = new FolderPicker { InputPath = settingsModel.RootDrive };

            if (dialog.ShowDialog(Handle) == true)
            {
                settingsModel.RootDrive = dialog.InputPath;
            }
        }

        /// <summary>
        /// Processes the project ID entered by the user.
        /// </summary>
        /// <param name="projectID">The project ID entered by the user.</param>
        private void ProcessProjectID(string projectID)
        {
            if (!projectID.SafeProjectID())
            {
                MessageBox.Show(Resources.InvalidProjectIDMessage ?? "Invalid Project ID",
                                Resources.ErrorTitle ?? "Error",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Error);
                return;
            }

            try
            {
                // Load settings for the given project ID
                settingsModel = SettingsHelpers.LoadProjectSettings(projectID);


                // Initialize paths and load data
                //itializePaths(); // Initialize paths for XML files
                LoadProjectData(); // Load project data from XML file .SaveAsPDF_Project.xml

                LoadEmployeeData(); // Load employee data from XML file .SaveAsPDF_Employees.xml

                UpdateUI(); // Update the UI elements based on the loaded data

                cmbSaveLocation.Text = settingsModel.DefaultSavePath;

                // Mark data as loaded
                _dataLoaded = true;
            }
            catch (FileNotFoundException ex)
            {
                MessageBox.Show($"A required file was not found: {ex.Message}",
                                "File Not Found",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Warning);
            }
            catch (UnauthorizedAccessException ex)
            {
                MessageBox.Show($"Access to a file or directory was denied: {ex.Message}",
                                "Access Denied",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Error);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An unexpected error occurred while processing the project ID: {ex.Message}",
                                "Error",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Loads the project data from the XML file.
        /// </summary>
        private void LoadProjectData()
        {
            if (File.Exists(settingsModel.XmlProjectFile))
            {
                try
                {
                    // Deserialize the XML file into a ProjectModel using XmlSerializer
                    _projectModel = XmlFileHelper.XmlProjectFileToModel(settingsModel.XmlProjectFile);

                    //XmlSerializer serializer = new XmlSerializer(typeof(ProjectModel));
                    //using (FileStream fileStream = new FileStream(settingsModel.XmlProjectFile, FileMode.Open))
                    //{
                    //    _projectModel = (ProjectModel)serializer.Deserialize(fileStream);
                    //}

                    // Update the UI with the loaded project data
                    if (_projectModel != null)
                    {
                        txtProjectName.Text = _projectModel.ProjectName;
                        chkbSendNote.Checked = _projectModel.NoteToProjectLeader;
                        rtxtProjectNotes.Text = _projectModel.ProjectNotes;
                    }
                }
                catch (Exception ex)
                {
                    // Handle XML file error
                    MessageBox.Show($"Error loading project data: {ex.Message}\nThe file will be overwritten with default values.",
                                    "SaveAsPDF:LoadProjectData",
                                    MessageBoxButtons.OK,
                                    MessageBoxIcon.Warning);

                    // Overwrite the XML file with default values
                    _projectModel = new ProjectModel
                    {
                        ProjectName = "Default Project",
                        ProjectNumber = "0000",
                        NoteToProjectLeader = false,
                        DefaultSaveFolder = settingsModel.DefaultSavePath,
                        ProjectNotes = "Default notes",
                        LastSavePath = settingsModel.DefaultSavePath
                    };

                    // Save the default project model to the XML file
                    settingsModel.XmlProjectFile.ProjectModelToXmlFile(_projectModel);

                    // Update the UI with default values
                    txtProjectName.Text = _projectModel.ProjectName;
                    chkbSendNote.Checked = _projectModel.NoteToProjectLeader;
                    rtxtProjectNotes.Text = _projectModel.ProjectNotes;
                }
            }
            else
            {
                // If the file does not exist, create it with default values
                _projectModel = new ProjectModel
                {
                    ProjectName = "Default Project",
                    ProjectNumber = "0000",
                    NoteToProjectLeader = false,
                    DefaultSaveFolder = settingsModel.DefaultSavePath,
                    ProjectNotes = "Default notes",
                    LastSavePath = settingsModel.DefaultSavePath
                };

                // Save the default project model to the XML file
                settingsModel.XmlProjectFile.ProjectModelToXmlFile(_projectModel);

                // Update the UI with default values
                txtProjectName.Text = _projectModel.ProjectName;
                chkbSendNote.Checked = _projectModel.NoteToProjectLeader;
                rtxtProjectNotes.Text = _projectModel.ProjectNotes;
            }
        }


        /// <summary>
        /// Loads the employee data from the XML file.
        /// </summary>
        private void LoadEmployeeData()
        {
            dgvEmployees.Rows.Clear();
            if (File.Exists(settingsModel.XmlProjectFile))
            {
                _employeesModel = settingsModel.XmlEmployeesFile.XmlEmployeesFileToModel();
                if (_employeesModel != null)
                {
                    foreach (EmployeeModel em in _employeesModel)
                    {
                        dgvEmployees.Rows.Add(em.Id, em.FirstName, em.LastName, em.EmailAddress);
                    }
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
                //Clear and reload the save location combo box

                cmbSaveLocation.Items.Clear();
                cmbSaveLocation.LoadComboBoxWithPaths(settingsModel.DefaultTreeFile, txtProjectID.Text);
                cmbSaveLocation.CustomizeComboBox(); //highlight the paths exist on the template but not exist on the file system

                // Set the default selected index
                if (cmbSaveLocation.Items.Count > 0)
                {
                    cmbSaveLocation.SelectedIndex = 0;
                }

                // Clear and reload the folder tree view
                tvFolders.Nodes.Clear();
                if (settingsModel.ProjectRootFolder.Exists)
                {
                    tvFolders.Nodes.Add(TreeHelpers.CreateDirectoryNode(settingsModel.ProjectRootFolder));
                    tvFolders.ExpandAll();
                    tvFolders.SelectedNode = tvFolders.Nodes[0];
                }
                else
                {
                    MessageBox.Show("The project root folder does not exist.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }

                // Update the full path text box
                txtFullPath.Text = settingsModel.ProjectRootFolder.FullName;


                // Set focus to the OK button
                btnOK.Focus();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An error occurred while updating the UI: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        /// <summary>
        /// Process the mail item and display its subject and attachments in the UI.
        /// </summary>
        /// <param name="mailItem"></param>
        private void ProcessMailItem(MailItem mailItem)
        {
            txtSubject.Text = mailItem.Subject;

            var attachments = mailItem.GetAttachmentsFromEmail();
            int i = 0;
            foreach (var attachment in attachments)
            {
                if (attachment != null)
                {
                    var attachmentsModel = new AttachmentsModel
                    {
                        attachmentId = i,
                        isChecked = true,
                        fileName = attachment.FileName,
                        fileSize = attachment.Size.BytesToString()
                    };
                    i++;
                    attachmentsModels.Add(attachmentsModel);
                }
            }

            //load to the dataGridView 
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
        /// Checks if the selected item is a mail item.
        /// </summary>
        private void ShowInvalidMailItemError()
        {
            MessageBox.Show("יש לבחור הודעות דואר אלקטרוני בלבד", "SaveAsPDF");
            Close();
        }


        /// <summary>
        /// clear the form from previews use
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
        private void btnOK_Click(object sender, EventArgs e)
        {
            //if (!_dataLoaded)
            //{
            //    LoadXmls();
            //}

            #region Populate Models

            //update and save settings model 
            settingsModel.OpenPDF = chbOpenPDF.Checked;

            //SettingsHelpers.SaveModelToSettings(settingsModel);

            //build _projectModel model
            _projectModel.ProjectName = txtProjectName.Text;
            _projectModel.ProjectNumber = txtProjectID.Text; //not in use
            _projectModel.NoteToProjectLeader = chkbSendNote.Checked;
            _projectModel.ProjectNotes = rtxtProjectNotes.Text;
            _projectModel.LastSavePath = cmbSaveLocation.SelectedText;

            //build the Employees model
            _employeesModel = dgvEmployees.DgvEmployeesToModel();
            #endregion

            #region Create XML files for the models

            //create _projectModel XML file
            settingsModel.XmlProjectFile.ProjectModelToXmlFile(_projectModel);

            //create the _employeesModel XML file from List<EmployeeModel> 
            settingsModel.XmlEmployeesFile.EmployeesModelToXmlFile(_employeesModel);

            #endregion
            string sPath = cmbSaveLocation.Text; //the current selected path in the combo box
            if (!string.IsNullOrEmpty(sPath))
            {
                DirectoryInfo directory = new DirectoryInfo(sPath);
                if (!directory.Exists)
                {
                    FileFoldersHelper.CreateDirectory(directory.FullName);
                }
            }
            else
            {
                var Dialog = new FolderPicker();
                Dialog.InputPath = settingsModel.RootDrive;
                if (Dialog.ShowDialog(Handle) == true)
                {
                    sPath = Dialog.ResultPath;
                }
            }
            //save the _mailItem to the current working directory and create an attachment list to add to PDF/mail   

            List<string> attachmentsList = new List<string>();
            //convert the message to HTML 
            if (_mailItem.BodyFormat != OlBodyFormat.olFormatHTML)
            {
                _mailItem.BodyFormat = OlBodyFormat.olFormatHTML;
            }

            // Save the attachments and retrieve their filenames
            attachmentsList.AddRange(_mailItem.SaveAttachments(dgvAttachments, sPath, false));

            // Build the HTML table style and header
            string tableStyle = @"
<html>
<head>
    <style>
        table, td, th {
            border-collapse: collapse;
        }
        tr:nth-child(even) {
            background-color: #f2f2f2;
        }
    </style>
</head>
<body>
    <table style='float:right;'>
        <tr><td colspan='3'></td></tr>
        <tr style='text-align:center'>
            <th colspan='3'>SaveAsPDF ver." + Assembly.GetExecutingAssembly().GetName().Version + @"</th>
        </tr>
        <tr><td colspan='3'></td></tr>
        <tr style='text-align:right'>
            <td colspan='3'><a href='file://" + sPath + @"'>" + sPath + @"</a> :ההודעה נשמרה ב</td>
        </tr>
        <tr style='text-align:right'>
            <td colspan='3'>" + DateTime.Now.ToString("HH:mm dd/MM/yyyy") + @" :תאריך שמירה </td>
        </tr>
        <tr><td colspan='3'></td></tr>";

            // Add project-related metadata
            string projectData = $@"
<tr style='text-align:right'><td></td><td>{txtProjectName.Text}</td><th>שם הפרויקט</th></tr>
<tr style='text-align:right'><td></td><td>{txtProjectID.Text}</td><th>מס' פרויקט</th></tr>
<tr style='text-align:right'><td></td><td>{rtxtNotes.Text.Replace(Environment.NewLine, "<br>")}</td><th>הערות</th></tr>
<tr style='text-align:right'><td></td><td>{Environment.UserName}</td><th>שם משתמש</th></tr>";

            // Convert attachments and employee info to HTML rows
            string attachmetsString = attachmentsList.AttachmentsToString(sPath);
            string employeeString = dgvEmployees.dgvEmployeesToString();

            // Append all sections and close HTML structure
            _mailItem.HTMLBody = tableStyle + projectData + employeeString + attachmetsString + "</table></body>" + _mailItem.HTMLBody;

            // Save mail as PDF and persist changes
            _mailItem.SaveToPDF(sPath);
            _mailItem.Save();
            Close();

            // Open the PDF file if the checkbox is checked
            if (chbOpenPDF.Checked)
            {
                string pdfFilePath = Path.Combine(sPath, $"{_mailItem.Subject}.pdf");
                if (File.Exists(pdfFilePath))
                {
                    System.Diagnostics.Process.Start(pdfFilePath);
                }
                else
                {
                    MessageBox.Show("The PDF file could not be found.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            {
                //TODO1:0 Open the PDF file

                MessageBox.Show("open PDF =" + chbOpenPDF.Checked.ToString());
            }

            //_settingsModel.Save();
            //TODO: save settings 

        }

        private void BtnSettings_Click(object sender, EventArgs e)
        {
            FormSettings frm = new FormSettings(this);
            frm.ShowDialog();
        }

        private void dgvAttachments_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            //TODO1: Open attachment when double clicking it on the list
            //1. save attachment to temp folder 
            //string tmpFoder = @System.IO.Path.GetTempPath();
            //2. exec. the file using default file association 
            MessageBox.Show(dgvAttachments.CurrentCell.Value.ToString());
        }

        private void btnStyle_Click(object sender, EventArgs e)
        {

            if (dlgFont.ShowDialog() == DialogResult.OK)
            {
                rtxtNotes.SelectionFont = dlgFont.Font;
            }

        }

        private void btnPhoneBook_Click(object sender, EventArgs e)
        {
            FormContacts frmContacts = new FormContacts(this);
            frmContacts.ShowDialog(this);
        }

        private void chkbSendNote_CheckedChanged(object sender, EventArgs e)
        {
            foreach (EmployeeModel employee in _employeesModel)
            {
                SendEmailToEmployee(employee.EmailAddress);
            }

        }
        /// <summary>
        /// after the message was saved, send an email to the lead engineer about the saved message 
        /// </summary>
        /// <param name="EmailAddress"></param>
        private void SendEmailToEmployee(string EmailAddress) =>
            //TODO3: send notification email message to employee - command visible = 0 (Next Version?)
            MessageBox.Show($"Send email to {EmailAddress}");

        private void chkbSelectAllAttachments_CheckedChanged(object sender, EventArgs e)
        {
            if (!chkbSelectAllAttachments.Checked)
            {
                chkbSelectAllAttachments.Text = "בחר הכל";
            }
            else
            {
                chkbSelectAllAttachments.Text = "הסר הכל";
            }

            if (dgvAttachments.RowCount != 0)
            {

                foreach (DataGridViewRow row in dgvAttachments.Rows)
                {
                    row.Cells[1].Value = chkbSelectAllAttachments.Checked;

                }
            }
        }
        /// <summary>
        /// Load the settings model from settings.settings to pass to Settings form 
        /// settings form is the interface for settings model 
        /// </summary>
        /// <param name="settings"></param>
        public void SettingsComplete(SettingsModel settings)
        {
            settingsModel = SettingsHelpers.LoadSettingsToModel(settings);

        }


        /// <summary>
        /// Load the employee model from the data-grid and pass it to contacts form 
        /// </summary>
        /// <param name="model"></param>
        public void EmployeeComplete(EmployeeModel model)
        {
            bool found = false;

            foreach (DataGridViewRow row in dgvEmployees.Rows)
            {
                if (row.Cells[3].Value.ToString() == model.EmailAddress)
                {
                    found = true;
                }
            }

            if (!found)
            {
                //add new employee(s) to the list 
                _employeesModel.Add(model);
                dgvEmployees.Rows.Add(model.Id.ToString(),
                                        model.FirstName,
                                        model.LastName,
                                        model.EmailAddress);
            }

        }

        private void tvFolders_BeforeExpand(object sender, TreeViewCancelEventArgs e)
        {
            //suppress the tree expand on double click
            if (_isDoubleClick && e.Action == TreeViewAction.Expand)
            {
                e.Cancel = true;
            }

        }


        private void RemoveEmployee_Click(object sender, EventArgs e)
        {
            //if (e.KeyCode == Keys.Delete)
            //{
            Int32 selectedRowCount = dgvEmployees.Rows.GetRowCount(DataGridViewElementStates.Selected);
            if (selectedRowCount > 0)
            {
                for (int i = 0; i < selectedRowCount; i++)
                {
                    dgvEmployees.Rows.RemoveAt(dgvEmployees.SelectedRows[0].Index);
                }
            }
            //}

        }

        private void btnNewProject_Click(object sender, EventArgs e)
        {
            FormNewProject frm = new FormNewProject(this);
            frm.ShowDialog(this);
        }

        public void NewProjectComplete(ProjectModel model)
        {
            _projectModel = model;

            txtProjectID.Text = _projectModel.ProjectNumber;
            txtProjectName.Text = _projectModel.ProjectName;
            rtxtProjectNotes.Text = _projectModel.ProjectNotes;
            //TODO1: refresh folder tree-view

        }

        private void txtProjectID_MouseHover(object sender, EventArgs e)
        {
            tsslStatus.Enabled = true;
            tsslStatus.Text = "מספר פרויקט כפי שמופיע במסטרפלן";
        }

        private void btnCopyNotesToMail_Click(object sender, EventArgs e)
        {
            rtxtNotes.Text += $"\n {rtxtProjectNotes.Text}";
        }

        private void btnCopyNotesToProject_Click(object sender, EventArgs e)
        {
            rtxtProjectNotes.Text += $"\n {rtxtNotes.Text}";
        }

        private void tvFolders_AfterSelect(object sender, TreeViewEventArgs e)
        {
            TreeNode CurrentNode = e.Node;
            string fullpath = CurrentNode.FullPath;
            mySelectedNode = CurrentNode;

            //txtFullPath.Text = $@"{settingsModel.ProjectRootFolder.FullName.Trim('\\')}{settingsModel.ProjectRootFolder.ToString().Replace(
            //                         settingsModel.ProjectRootTag, string.Empty)}";// no need the '\'
        }


        private void tvFolders_AfterLabelEdit(object sender, NodeLabelEditEventArgs e)
        {
            if (e.Label == null)
            {
                // Cancel the edit if the label is null (user pressed ESC or left it empty)
                e.CancelEdit = true;
                MessageBox.Show("שם לא חוקי.\n לא ניתן ליצור שם ריק. חובה תו אחד לפחות", "עריכת שם");
                return;
            }

            string nodeNewLabel = e.Label.SafeFolderName();

            if (string.IsNullOrWhiteSpace(nodeNewLabel))
            {
                // Cancel the edit if the label is empty or whitespace
                e.CancelEdit = true;
                MessageBox.Show("שם לא חוקי.\n לא ניתן ליצור שם ריק. חובה תו אחד לפחות", "עריכת שם");
                return;
            }

            if (nodeNewLabel.IndexOfAny(new char[] { '\\', '/', ':', '*', '?', '<', '>', '|', '"' }) != -1)
            {
                // Cancel the edit if the label contains invalid characters
                e.CancelEdit = true;
                MessageBox.Show("שם לא חוקי.\nאין להשתמש בתווים הבאים \n'\\', '/', ':', '*', '?', '<', '>', '|' '\"' ", "עריכת שם");
                return;
            }

            try
            {
                // Rename the directory
                string oldPath = Path.Combine(settingsModel.ProjectRootFolder.Parent.FullName, e.Node.FullPath);
                string newPath = Path.Combine(settingsModel.ProjectRootFolder.Parent.FullName, e.Node.Parent.FullPath, nodeNewLabel);

                DirectoryInfo directoryInfo = new DirectoryInfo(oldPath);
                directoryInfo.RenameDirectory(newPath);

                // Update the TreeView node
                e.Node.Text = nodeNewLabel;
            }
            catch (Exception ex)
            {
                // Handle any errors during the renaming process
                e.CancelEdit = true;
                MessageBox.Show($"Error renaming folder: {ex.Message}\n{Path.Combine(settingsModel.ProjectRootFolder.Parent.FullName, e.Node.FullPath)}", "SaveAsPDF:tvFolders_AfterLabelEdit");
            }
        }



        private void chbOpenPDF_CheckedChanged(object sender, EventArgs e)
        {
            //TODO1: chbOpenPDF_CheckedChanged:make sure it works 

            //Settings.Default.OpenPDF = chbOpenPDF.Checked;
            //Settings.Default.Save();

            settingsModel.OpenPDF = chbOpenPDF.Checked;


        }

        private void tvFolders_NodeMouseDoubleClick(object sender, TreeNodeMouseClickEventArgs e)
        {
            //TODO1: open the folder in explorer and update the path in the cmbSaveLocation

            string path = $@"{settingsModel.ProjectRootFolder.Parent.FullName}\{e.Node.FullPath}";
            System.Diagnostics.Process.Start("explorer.exe", path);

            cmbSaveLocation.SelectedText = path;


        }

        private void tvFolders_NodeMouseClick(object sender, TreeNodeMouseClickEventArgs e)
        {

            string path = $@"{settingsModel.ProjectRootFolder.Parent.FullName}\{e.Node.FullPath}";



            //update cmbSaveLocation with the selected node path
            //cmbSaveLocation.Items.Clear();

            //clear the selected item in cmbSaveLocation 
            cmbSaveLocation.Select();
            cmbSaveLocation.SelectedItem = null;
            //enter the selected path to the combo box
            cmbSaveLocation.SelectedText = path;

        }



        private void txtProjectID_Validating(object sender, System.ComponentModel.CancelEventArgs e)
        {
            if (!txtProjectID.Text.SafeProjectID())
            {
                // Cancel the event and select the text to be corrected by the user.
                errorProviderMain.SetError(txtProjectID, "מספר פרויקט לא חוקי");
                txtProjectID.Select(0, txtProjectID.Text.Length);
                tsslStatus.Text = errorProviderMain.GetError(txtProjectID);
                e.Cancel = true;
            }
        }

        /// <summary>
        /// Handles the Validated event for the txtProjectID TextBox.
        /// This method is triggered after the txtProjectID control has been successfully validated.
        /// It performs the following actions:
        /// - Clears any error messages associated with the txtProjectID control.
        /// - Updates the status label with the current error state of txtProjectID.
        /// - Processes the entered project ID by calling <see cref="ProcessProjectID(string)"/>.
        /// - Updates the autocomplete source with the entered project ID.
        /// - Handles the first run scenario if the root drive is not set in the settings.
        /// - Processes the current mail item if it is valid; otherwise, displays an error message.
        /// - Sets the data loaded flag to true.
        /// </summary>
        /// <param name="sender">The source of the event, typically the txtProjectID control.</param>
        /// <param name="e">An <see cref="EventArgs"/> that contains the event data.</param>
        private void txtProjectID_Validated(object sender, EventArgs e)
        {
            errorProviderMain.SetError(txtProjectID, string.Empty); // Clear the error
            tsslStatus.Text = errorProviderMain.GetError(txtProjectID);

            ProcessProjectID(txtProjectID.Text);
            //XMessageBox.Show("הפרויקט נשמר בהצלחה", "SaveAsPDF", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
            XMessageBox.Show("הפרויקט נשמר בהצלחה", "SaveAsPDF", XMessageBoxButtons.OK, XMessageBoxIcon.Information);
            UpdateAutoCompleteSource(txtProjectID.Text);

            if (string.IsNullOrEmpty(settingsModel.RootDrive))
            {
                HandleFirstRun();
            }

            if (_mailItem is MailItem mailItem)
            {
                ProcessMailItem(_mailItem);
            }
            else
            {
                ShowInvalidMailItemError();
            }

            _dataLoaded = true;
        }

        private void FormMain_FormClosing(object sender, FormClosingEventArgs e)
        {
            // Close all open files

            //TODO: Implement code to close all open files
            //if (e.CloseReason == CloseReason.UserClosing)
            //{
            //    if (MessageBox.Show("האם לצאת מהיישום?", "SaveAsPDF", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
            //    {
            //        e.Cancel = true;
            //    }
            //}

            // Rest of the code

        }

        private void tvFolders_BeforeCollapse(object sender, TreeViewCancelEventArgs e)
        {
            if (_isDoubleClick && e.Action == TreeViewAction.Collapse)
            {
                e.Cancel = true;
            }
        }

        private void tvFolders_MouseDown(object sender, MouseEventArgs e)
        {
            _isDoubleClick = e.Clicks > 1;
        }

        private void FormMain_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.F2)
            {
                if (tvFolders.SelectedNode != null)
                {
                    tvFolders.SelectedNode.BeginEdit();
                }
            }
            else if (e.KeyCode == Keys.Delete)
            {
                if (tvFolders.SelectedNode != null)
                {
                    if (tvFolders.SelectedNode.Nodes.Count == 0)
                    {
                        try
                        {
                            DirectoryInfo directoryInfo = new DirectoryInfo(Path.Combine(settingsModel.ProjectRootFolder.Parent.FullName, tvFolders.SelectedNode.FullPath));
                            directoryInfo.Delete(true);
                            tvFolders.SelectedNode.Remove();
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(ex.Message, "SaveAsPDF:tvFolders_KeyDown");
                        }
                    }
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
                {
                    Close();
                }
                txtProjectID.Clear();
            }
        }

        private void cmbSaveLocation_SelectedValueChanged(object sender, EventArgs e)
        {
            // Ensure _projectModel is initialized before interacting with it
            if (_projectModel == null)
            {
                // Initialize _projectModel with default values if it is null
                //TODO: make sure defaults are correct
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

            // Save the selected path to the _projectModel model
            _projectModel.LastSavePath = cmbSaveLocation.SelectedItem?.ToString();
        }

        private void cmbSaveLocation_TextUpdate(object sender, EventArgs e)
        {
            // Update the breadcrumb path in the combo box "C > 12 > 123"
            cmbSaveLocation.SetBreadcrumbPath();
        }
    }
}