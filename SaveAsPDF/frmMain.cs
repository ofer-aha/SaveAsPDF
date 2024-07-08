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
    public partial class frmMain : Form, IEmployeeRequester, INewProjectRequester, ISettingsRequester
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

        public frmMain()
        {
            InitializeComponent();

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

                KeyDown += frmMain_KeyDown; //add the key-down event to the form
            }
        }

        private void LoadSearchHistory()
        {
            if (Settings.Default.LastProjects == null)
            {
                Settings.Default.LastProjects = new StringCollection();
            }
            txtProjectID.AutoCompleteCustomSource.AddRange(Settings.Default.LastProjects.Cast<string>().ToArray());
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
                txtProjectID_Validating(sender, new System.ComponentModel.CancelEventArgs());
            }
        }

        /// <summary>
        /// Updates the auto complete source with the current text in the project ID textbox.
        /// </summary>
        private void UpdateAutoCompleteSource()
        {
            if (!searchHistory.Contains(txtProjectID.Text))
            {
                searchHistory.Add(txtProjectID.Text);
                Settings.Default.LastProjects = new StringCollection();
                Settings.Default.LastProjects.AddRange(searchHistory.ToArray());
                Settings.Default.Save();
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
            //ClearForm();
            //settingsModel.ProjectRootFolder = txtProjectID.Text.ProjectFullPath(settingsModel.RootDrive);
            //txtSaveLocation.Text = $@"{settingsModel.ProjectRootFolder.FullName}
            //                  {settingsModel.DefaultSavePath.Replace(settingsModel.ProjectRootTag,
            //                          settingsModel.ProjectRootFolder.FullName).Replace(settingsModel.DateTag, DateTime.Now.ToString("dd.MM.yyyy"))}";
            if (projectID.SafeProjectID())
            {
                settingsModel = SettingsHelpers.LoadProjectSettings(projectID);
                //SettingsModel settings = new SettingsModel
                //{
                //    ProjectRootFolder = settingsModel.ProjectRootFolder,
                //};

                LoadXmls();

                txtFullPath.Text = settingsModel.ProjectRootFolder.FullName;
                txtSaveLocation.Text = settingsModel.DefaultSavePath;

                btnOK.Focus();
                _dataLoaded = true;
            }
        }
        private void ShowInvalidProjectIDError()
        {
            txtProjectID.BackColor = System.Drawing.Color.Red;
            txtProjectID.SelectAll();
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
        /// Load XML files from .SaveAsPDF folder to the models and construct the full path for everything
        /// </summary>
        private void LoadXmls()
        {
            if (settingsModel.ProjectRootFolder.Exists)
            {
                // j:\12\1245\  - exist 
                // construct the full path for everything
                settingsModel.ProjectRootFolder.FullName.CreateHiddenFolder();
                _xmlSaveAsPdfFolder = new DirectoryInfo(settingsModel.XmlSaveAsPDFFolder);
                _xmlProjectFile = settingsModel.XmlProjectFile;
                _xmlEmploeeysFile = settingsModel.XmlEmployeesFile;

                if (File.Exists(_xmlProjectFile))
                {
                    //load the XML file to _projectModel model
                    _projectModel = _xmlProjectFile.XmlProjectFileToModel();

                    if (_projectModel != null)
                    {
                        txtProjectName.Text = _projectModel.ProjectName;
                        chkbSendNote.Checked = _projectModel.NoteEmployee;
                        rtxtProjectNotes.Text = _projectModel.ProjectNotes;
                    }
                }

                dgvEmployees.Rows.Clear();
                //load the XML file to Employees list-box
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

            tvFolders.Nodes.Clear();
            tvFolders.Nodes.Add(TreeHelpers.CreateDirectoryNode(settingsModel.ProjectRootFolder));
            tvFolders.ExpandAll();
            tvFolders.SelectedNode = tvFolders.Nodes[0];
        }
        /// <summary>
        /// clear the form from previews use
        /// </summary>
        private void ClearForm()
        {
            txtProjectName.Clear();
            txtFullPath.Clear();
            txtSaveLocation.Clear();
            rtxtNotes.Clear();
            rtxtProjectNotes.Clear();
            dgvAttachments.DataSource = null;
            dgvEmployees.DataSource = null;
            tvFolders.DataBindings.Clear();

        }
        private void btnOK_Click(object sender, EventArgs e)
        {
            if (!_dataLoaded)
            {
                LoadXmls();
            }

            #region Populate Models

            //update and save settings model 
            settingsModel.OpenPDF = chbOpenPDF.Checked;

            //SettingsHelpers.SaveModelToSettings(settingsModel);

            //build _projectModel model
            _projectModel.ProjectName = txtProjectName.Text;
            _projectModel.ProjectNumber = txtProjectID.Text; //not in use
            _projectModel.NoteEmployee = chkbSendNote.Checked;
            _projectModel.ProjectNotes = rtxtProjectNotes.Text;

            //build the Employees model
            _employeesModel = dgvEmployees.DgvEmployeesToModel();
            #endregion

            #region Create XML files for the models

            //create the SaveAsPDF hidden folder
            //_xmlSaveAsPdfFolder.FullName.CreateHiddenFolder(); //already doing it on LoadXml() 

            //create _projectModel XML file
            _xmlProjectFile.ProjectModelToXmlFile(_projectModel);

            //create the _employeesModel XML file from List<EmployeeModel> 
            _xmlEmploeeysFile.EmployeesModelToXmlFile(_employeesModel);

            #endregion

            if (!string.IsNullOrEmpty(txtSaveLocation.Text))
            {
                DirectoryInfo directory = new DirectoryInfo(txtSaveLocation.Text);
                if (!directory.Exists)
                {
                    FileFoldersHelper.MkDir(directory.FullName);
                }
            }
            else
            {
                var Dialog = new FolderPicker();
                Dialog.InputPath = settingsModel.RootDrive;
                if (Dialog.ShowDialog(Handle) == true)
                {
                    txtSaveLocation.Text = Dialog.ResultPath;
                }
            }
            //save the _mailItem to the current working directory and create an attachment list to add to PDF/mail   

            List<string> attList = new List<string>();
            //convert the message to HTML 
            if (_mailItem.BodyFormat != OlBodyFormat.olFormatHTML)
            {
                _mailItem.BodyFormat = OlBodyFormat.olFormatHTML;
            }

            //Save the attachments and returning the actual file name list
            attList.AddRange(_mailItem.SaveAttachments(dgvAttachments, txtSaveLocation.Text, false));

            //string tableStyle = "<style> table, th,td {border: 1px solid black; text-align: right;}</style><br>";
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
                                $"<tr style=\"text-align:right\"><td colspan=\"3\"><a href='file://{txtSaveLocation.Text}'>{txtSaveLocation.Text}</a> :ההודעה נשמרה ב</td></tr>" +
                                $"<tr style=\"text-align:right\"><td colspan=\"3\">{DateTime.Now.ToString("HH:mm dd/MM/yyyy")} :תאריך שמירה </td></tr>" +
                                "<tr><td colspan=\"3\"></td></tr>"; //empty line 

            string projData = $"<tr style=\"text-align:right\"><td></td><td>{txtProjectName.Text}</td><th>שם הפרויקט</th></tr>" +
                              $"<tr style=\"text-align:right\"><td></td><td>{txtProjectID.Text}</td><th >מס' פרויקט</th></tr>" +
                              $"<tr style=\"text-align:right\"><td></td><td>{rtxtNotes.Text.Replace(Environment.NewLine, "<br>")}</td><th >הערות</th></tr>" +
                              $"<tr style=\"text-align:right\"><td></td><td>{Environment.UserName}</td><th>שם משתמש</th></tr>";


            string attString = attList.AttachmentsToString(txtSaveLocation.Text);
            string empString = dgvEmployees.dgvEmployeesToString();

            //construct the HTMLbody message 
            _mailItem.HTMLBody = tableStyle +
                                projData +
                                empString +
                                attString +
                                "</table></body>" +
                                _mailItem.HTMLBody;


            OfficeHelpers.SaveToPDF(_mailItem, txtSaveLocation.Text);


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
            frmSettings frm = new frmSettings(this);
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
            frmContacts frmContacts = new frmContacts(this);
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






            //tvFolders.Nodes.Clear();
            // tvFolders.Nodes.Add(TreeHelper.TraverseDirectory(e.Node.Text, 1));
            //if (e.Node.Nodes.Count > 0)
            //{
            //    if (e.Node.Nodes[0].Text == "..." && e.Node.Nodes[0].Tag == null)
            //    {
            //        e.Node.Nodes.Clear();

            //        //get the list of sub directory
            //        string[] dirs = Directory.GetDirectories(e.Node.Tag.ToString());

            //        foreach (string dir in dirs)
            //        {
            //            DirectoryInfo di = new DirectoryInfo(dir);
            //            TreeNode node = new TreeNode(di.Name, 0, 1);

            //            try
            //            {
            //                //keep the directory's full path in the tag for use later
            //                node.Tag = dir;

            //                //if the directory has sub directories add the place holder
            //                if (di.GetDirectories().Count() > 0)
            //                    node.Nodes.Add(null, "...", 0, 0);
            //            }
            //            catch (UnauthorizedAccessException)
            //            {
            //                //display a locked folder icon
            //                node.ImageIndex = 12;
            //                node.SelectedImageIndex = 12;
            //            }
            //            catch (Exception ex)
            //            {
            //                MessageBox.Show(ex.Message, "SaveAsPDF: treeView1_BeforeExpand",
            //                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            //            }
            //            finally
            //            {
            //                e.Node.Nodes.Add(node);
            //            }
            //        }
            //    }
            //}
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
            frmNewProject frm = new frmNewProject();
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
                    directoryInfo.RnDir($@"{settingsModel.ProjectRootFolder.Parent.FullName}\{e.Node.Parent.FullPath}\{nodeNewLable}"); //nodeNewLable = new SAFE name

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
            //TODO1: make sure it works 
            //_settingsModel.OpenPDF = chbOpenPDF.Checked;
            Settings.Default.OpenPDF = chbOpenPDF.Checked;
            Settings.Default.Save();

        }

        private void tvFolders_NodeMouseDoubleClick(object sender, TreeNodeMouseClickEventArgs e)
        {
            txtSaveLocation.Text = $@"{settingsModel.ProjectRootFolder.Parent.FullName}\{e.Node.FullPath}";

            //try
            //{
            //    if (Directory.Exists(e.Node.FullPath))
            //    {
            //        Process.Start(e.Node.FullPath);
            //    }
            //    else
            //    {
            //        Process.Start($@"{settingsModel.RootDrive}\{e.Node.FullPath}");
            //    }
            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show($"ex:{ex.Message}\n settingsmodel:{settingsModel.ProjectRootFolder}\n e.node:{e.Node.FullPath}");
            //}
        }

        private void tvFolders_NodeMouseClick(object sender, TreeNodeMouseClickEventArgs e)
        {
            //_mySelectedNode = e.Node;
            //MessageBox.Show($"e.node.Name:{e.Node.Text} _mySelectedNode: {_mySelectedNode.FullPath}");
        }



        private void txtProjectID_Validating(object sender, System.ComponentModel.CancelEventArgs e)
        {
            if (!txtProjectID.Text.SafeProjectID())
            {
                e.Cancel = true;
                errorProviderMain.SetError(txtProjectID, "מספר פרויקט לא חוקי");
                txtProjectID.Select(0, txtProjectID.Text.Length);
                txtProjectID.BackColor = System.Drawing.Color.Red;
                tsslStatus.Text = errorProviderMain.GetError(txtProjectID);
                ShowInvalidProjectIDError();
            }
        }

        private void txtProjectID_Validated(object sender, EventArgs e)
        {
            errorProviderMain.SetError(txtProjectID, string.Empty);
            txtProjectID.BackColor = System.Drawing.Color.White;
            tsslStatus.Text = errorProviderMain.GetError(txtProjectID);
            //settingsModel = SettingsHelpers.LoadProjectSettings(((TextBox)sender).Text);
            //SettingsModel settings = new SettingsModel
            //{
            //    ProjectRootFolder = settingsModel.ProjectRootFolder,
            //};
            ProcessProjectID(txtProjectID.Text);
            UpdateAutoCompleteSource();


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

        private void frmMain_FormClosing(object sender, FormClosingEventArgs e)
        {
            // Close all open files
            // TODO: Implement code to close all open files

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

        private void frmMain_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.F2)
            {
                if (tvFolders.SelectedNode != null)
                {
                    tvFolders.SelectedNode.BeginEdit();
                }
            }
            if (e.KeyCode == Keys.Delete)
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
            if (e.KeyCode == Keys.F5)
            {
                tvFolders.Nodes.Clear();
                tvFolders.Nodes.Add(TreeHelpers.TraverseDirectory(settingsModel.ProjectRootFolder.FullName, 1));
            }
            if (e.KeyCode == Keys.Escape)
            {
                Close();
            }
        }
    }
}