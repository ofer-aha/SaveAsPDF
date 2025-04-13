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
        private DirectoryInfo _xmlSaveAsPdfFolder;
        private string _xmlProjectFile;
        private string _xmlEmploeeysFile;

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

        public static object settingmodel { get; internal set; }

        public FormMain()
        {
            InitializeComponent();
            this.Load += new EventHandler(FormMain_Load);
        }
        private string LoadEmailSubject()
        {
            // load the email subject from mail item
            if (_mailItem != null)
                return _mailItem.Subject;
            else
                return "Default Email Subject";

        }
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

        private void btnCancel_Click(object sender, EventArgs e)
        {
            //close the form - do  nothing
            Close();
        }
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
                InitializePaths();
                LoadProjectData();
                LoadEmployeeData();

                // Update the UI with the loaded data
                UpdateUI();

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

        private void InitializePaths()
        {
            _xmlSaveAsPdfFolder = new DirectoryInfo(settingsModel.XmlSaveAsPDFFolder);
            _xmlProjectFile = settingsModel.XmlProjectFile;
            _xmlEmploeeysFile = settingsModel.XmlEmployeesFile;
            FileFoldersHelper.CreateHiddenDirectory(_xmlSaveAsPdfFolder.FullName);
        }

        private void LoadProjectData()
        {
            if (File.Exists(_xmlProjectFile))
            {
                _projectModel = _xmlProjectFile.XmlProjectFileToModel();
                if (_projectModel != null)
                {
                    txtProjectName.Text = _projectModel.ProjectName;
                    chkbSendNote.Checked = _projectModel.NoteToProjectLeader;
                    rtxtProjectNotes.Text = _projectModel.ProjectNotes;
                }
            }
        }

        private void LoadEmployeeData()
        {
            dgvEmployees.Rows.Clear();
            if (File.Exists(_xmlEmploeeysFile))
            {
                _employeesModel = _xmlEmploeeysFile.XmlEmployeesFileToModel();
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
                // Clear and reload the save location combo box
                //cmbSaveLocation.Items.Clear();
                //cmbSaveLocation.LoadFromFile(settingsModel.DefaultTreeFile, txtProjectID.Text);

                cmbSaveLocation.Items.Clear();
                cmbSaveLocation.LoadFromFile(settingsModel.DefaultTreeFile, settingsModel.ProjectRootTag);
                //cmbSaveLocation.SetBreadcrumbPath(settingsModel.ProjectRootFolder.FullName);


                for (int i = 0; i < cmbSaveLocation.Items.Count; i++)
                {
                    if (cmbSaveLocation.Items[i] is string item)
                    {
                        cmbSaveLocation.Items[i] = item.Replace(settingsModel.ProjectRootTag,
                                                                settingsModel.ProjectRootFolder.FullName.TrimEnd('\\'));
                        //replace the DateTag with the current date
                        //cmbSaveLocation.Items[i] = item.Replace(settingsModel.DateTag,DateTime.Now.ToString("yyyy.MM.dd"));
                    }
                    else
                    {
                        throw new InvalidCastException($"Item at index {i} in cmbSaveLocation is not a string.");
                    }
                }

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

                // Handle the default save path in the combo box
                if (!string.IsNullOrEmpty(settingsModel.DefaultSavePath))
                {
                    if (cmbSaveLocation.Items.Contains(settingsModel.DefaultSavePath))
                    {
                        cmbSaveLocation.SelectedItem = settingsModel.DefaultSavePath;
                    }
                    else
                    {
                        cmbSaveLocation.Items.Add(settingsModel.DefaultSavePath);
                        cmbSaveLocation.SelectedItem = settingsModel.DefaultSavePath;
                    }
                }

                // Set focus to the OK button
                btnOK.Focus();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An error occurred while updating the UI: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

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
            _xmlProjectFile.ProjectModelToXmlFile(_projectModel);

            //create the _employeesModel XML file from List<EmployeeModel> 
            _xmlEmploeeysFile.EmployeesModelToXmlFile(_employeesModel);

            #endregion

            if (!string.IsNullOrEmpty(cmbSaveLocation.SelectedText))
            {
                DirectoryInfo directory = new DirectoryInfo(cmbSaveLocation.SelectedText);
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
                    cmbSaveLocation.SelectedText = Dialog.ResultPath;
                }
            }
            //save the _mailItem to the current working directory and create an attachment list to add to PDF/mail   

            List<string> attachmentsList = new List<string>();
            //convert the message to HTML 
            if (_mailItem.BodyFormat != OlBodyFormat.olFormatHTML)
            {
                _mailItem.BodyFormat = OlBodyFormat.olFormatHTML;
            }

            //Save the attachments and returning the actual file name list
            attachmentsList.AddRange(_mailItem.SaveAttachments(dgvAttachments, cmbSaveLocation.SelectedText, false));

            string tableStyle = "<html><head><style>" +
                                "table, td, th { " +
                                //"border: 0px solid blue; " +
                                //"border-collapse:collapse " +
                                "tr:nth-child(even) {background-color:#f2f2f2};" +
                                "}</style></head>" +
                                "<body>" +
                                "<table style=\"float:right;\">" +
                                "<tr><td colspan=\"3\"></td></tr>" + //empty line 
                                $"<tr style=\"text-align:center\"><th colspan=\"3\">SaveAsPDF ver.{Assembly.GetExecutingAssembly().GetName().Version.ToString()}</th></tr>" +
                                "<tr><td colspan=\"3\"></td></tr>" + //empty line 
                                $"<tr style=\"text-align:right\"><td colspan=\"3\"><a href='file://{cmbSaveLocation.SelectedText}'>{cmbSaveLocation.SelectedText}</a> :ההודעה נשמרה ב</td></tr>" +
                                $"<tr style=\"text-align:right\"><td colspan=\"3\">{DateTime.Now.ToString("HH:mm dd/MM/yyyy")} :תאריך שמירה </td></tr>" +
                                "<tr><td colspan=\"3\"></td></tr>"; //empty line 

            string projectData = $"<tr style=\"text-align:right\"><td></td><td>{txtProjectName.Text}</td><th>שם הפרויקט</th></tr>" +
                              $"<tr style=\"text-align:right\"><td></td><td>{txtProjectID.Text}</td><th >מס' פרויקט</th></tr>" +
                              $"<tr style=\"text-align:right\"><td></td><td>{rtxtNotes.Text.Replace(Environment.NewLine, "<br>")}</td><th >הערות</th></tr>" +
                              $"<tr style=\"text-align:right\"><td></td><td>{Environment.UserName}</td><th>שם משתמש</th></tr>";


            string attachmetsString = attachmentsList.AttachmentsToString(cmbSaveLocation.SelectedText);
            string employeeString = dgvEmployees.dgvEmployeesToString();

            //construct the HTMLbody message 
            _mailItem.HTMLBody = tableStyle +
                                projectData +
                                employeeString +
                                attachmetsString +
                                "</table></body>" +
                                _mailItem.HTMLBody;

            OfficeHelpers.SaveToPDF(_mailItem, cmbSaveLocation.SelectedText);

            Close();

            if (chbOpenPDF.Checked)
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
            string nodeNewLable = e.Label.SafeFolderName();

            if (nodeNewLable != null)
            {
                if (nodeNewLable.Length > 0)
                {

                    if (nodeNewLable.IndexOfAny(new char[] { '\\', '/', ':', '*', '?', '<', '>', '|', '"' }) == -1)

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
                           "אין להשתמש בתווים הבאים \n'\\', '/', ':', '*', '?', '<', '>', '|' '\"' ",
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
                    return;
                }

                try
                {

                    DirectoryInfo directoryInfo = new DirectoryInfo(Path.Combine(settingsModel.ProjectRootFolder.Parent.FullName, e.Node.FullPath));  // old path
                    directoryInfo.RenameDirectory($@"{settingsModel.ProjectRootFolder.Parent.FullName}\{e.Node.Parent.FullPath}\{nodeNewLable}"); //nodeNewLable = new SAFE name

                    string specificPath = $@"{e.Node.Parent.FullPath}\{e.Label}";

                    TreeNode existingNode = TreeHelpers.FindNodeByPath(tvFolders.Nodes, e.Node.FullPath);
                    if (existingNode != null)
                    {
                        TreeNode newNode = new TreeNode(nodeNewLable);
                        existingNode.Parent.Nodes.Insert(existingNode.Index, newNode.Text);
                        existingNode.Remove();

                    }
                }
                catch (Exception ex)
                {

                    MessageBox.Show(ex.Message + "\n" + Path.Combine(settingsModel.ProjectRootFolder.Parent.FullName, e.Node.FullPath), "SaveAsPDF:tvFolders_AfterLabelEdit");
                }

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
            //cmbSaveLocation.SelectedText = $@"{settingsModel.ProjectRootFolder.Parent.FullName}\{e.Node.FullPath}";
            cmbSaveLocation.SelectedText = path;

            //cmbSaveLocation.SelectedText = $@"{settingsModel.ProjectRootFolder.Parent.FullName}\{e.Node.FullPath}";

        }

        private void tvFolders_NodeMouseClick(object sender, TreeNodeMouseClickEventArgs e)
        {
            //update cmbSaveLocation with the selected node path
            cmbSaveLocation.SelectedText = $@"{settingsModel.ProjectRootFolder.Parent.FullName}\{e.Node.FullPath}";


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
            //save the selected path to the _projectModel model
            _projectModel.LastSavePath = cmbSaveLocation.SelectedText;

        }
    }
}