// Ignore Spelling: frm מפשחה הכל יש לבחור הודעות דואר אלקטרוני בלבד אימייל הסר הכול שם קובץ גודל יש לבחור הודעות דואר אלקטרוני בלבד ההודעה נשמרה ב  תאריך  שמירה  שם הפרויקט  מס פרויקט  הערות  שם משתמש בחר  הסר  מספר פרויקט כפי שמופיע במסטרפלן שם לא חוקי  אין להשתמש בתווים הבאים  עריכת שם שם לא חוקי לא ניתן ליצור שם ריק חובה תו אחד לפחות עריכת שם מספר פרויקט לא חוקי


using Microsoft.Office.Interop.Outlook;
using SaveAsPDF.Helpers;
using SaveAsPDF.Models;
using SaveAsPDF.Properties;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;
using Exception = System.Exception;

namespace SaveAsPDF
{
    public partial class frmMain : Form, IEmployeeRequester, INewProjectRequester, ISettingsRequester
    {
        private List<EmployeeModel> employeesModel = new List<EmployeeModel>();
        private ProjectModel projectModel = new ProjectModel();
        private SettingsModel settingsModel = new SettingsModel();


        // construct the full path for everything
        //public static DirectoryInfo sPath;
        private DirectoryInfo xmlSaveAsPdfFolder;
        private string xmlProjectFile;
        private string xmlEmploeeysFile;

        private bool dataLoaded = false;

        public static TreeNode mySelectedNode;

        private static MailItem mi = null;
        private MailItem mailItem = ThisAddIn.TypeOfMailitem(mi);

        DocumentModel oDoc = new DocumentModel();

        List<Attachment> attachments = new List<Attachment>();
        List<AttachmentsModel> attachmentsModels = new List<AttachmentsModel>();


        public frmMain()
        {
            InitializeComponent();

            //Load settings to settingsModel 
            SettingsHelpers.loadSettingsToModel();

            if (string.IsNullOrEmpty(settingsModel.RootDrive))
            {
                //TODO: handle first run 
                MessageBox.Show("this is the first run");
                var Dialog = new FolderPicker();
                Dialog.InputPath = settingsModel.RootDrive;

                if (Dialog.ShowDialog(Handle) == true)
                {
                    settingsModel.RootDrive = Dialog.InputPath;
                }
            }

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
        }
        private void frmMain_Load(object sender, EventArgs e)
        {
            settingsModel.ProjectRootFolders = txtProjectID.Text.ProjectFullPath();

            if (mailItem is MailItem && mailItem != null)
            {
                chkbSelectAllAttachments.Checked = true;
                chkbSelectAllAttachments.Text = "הסר הכל";
                txtSubject.Text = mailItem.Subject;
                txtSaveLocation.Text = settingsModel.ProjectRootFolders.FullName;

                //load the list from model

                attachments = mailItem.GetMailAttachments();
                int i = 0;
                foreach (Attachment attachment in attachments)
                {
                    if (attachment != null)
                    {
                        AttachmentsModel attachmentsModel = new AttachmentsModel();
                        attachmentsModel.attachmentId = i;
                        attachmentsModel.isChecked = true;
                        attachmentsModel.fileName = attachment.FileName;
                        attachmentsModel.fileSize = attachment.Size.BytesToString();
                        i++;
                        attachmentsModels.Add(attachmentsModel);
                    }

                }
                BindingSource source = new BindingSource();
                source.DataSource = attachmentsModels;
                dgvAttachments.DataSource = source;

                dgvAttachments.Columns[0].Visible = false;

                dgvAttachments.Columns[1].HeaderText = "V";
                dgvAttachments.Columns[1].ReadOnly = false;

                dgvAttachments.Columns[2].HeaderText = "שם קובץ";
                dgvAttachments.Columns[2].ReadOnly = true;
                dgvAttachments.Columns[3].HeaderText = "גודל";
                dgvAttachments.Columns[3].ReadOnly = true;
            }
            else
            {
                MessageBox.Show("יש לבחור הודעות דואר אלקטרוני בלבד", "SaveAsPDF");
                Close();
            }

            //get a list of the drives
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
            }
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            //close the form - do  nothing
            Close();
        }
        private void btnFolders_Click(object sender, EventArgs e)
        {
            var Dialog = new FolderPicker();

            if (!settingsModel.ProjectRootFolders.Exists)
            {
                Dialog.InputPath = settingsModel.RootDrive;
            }
            else
            {
                Dialog.InputPath = settingsModel.ProjectRootFolders.FullName;
            }

            if (Dialog.ShowDialog(Handle) == true)
            {
                tvFolders.Nodes.Clear();
                tvFolders.Nodes.Add(TreeHelper.TraverseDirectory(Dialog.ResultPath));
            }
        }
        private void txtProjectID_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                if (txtProjectID.Text.SafeProjectID())
                {
                    ClearForm();
                    LoadXmls();
                    btnOK.Focus();
                    dataLoaded = true;
                    return;
                }
                else
                {
                    picBoxProjectID.Visible = true;
                    txtProjectID.BackColor = System.Drawing.Color.Red;
                    txtProjectID.SelectAll();
                }
            }
        }
        private void LoadXmls()
        {
            // construct the full path for everything
            settingsModel.ProjectRootFolders = txtProjectID.Text.ProjectFullPath();
            settingsModel.DefaultSavePath = settingsModel.ProjectRootFolders + settingsModel.DefaultSavePath;
            //display the project's full path above the folder tree-view 
            txtFullPath.Text = settingsModel.ProjectRootFolders.ToString();

            xmlSaveAsPdfFolder = new DirectoryInfo(Path.Combine(settingsModel.ProjectRootFolders.FullName, settingsModel.XmlSaveAsPDFFolder));
            xmlProjectFile = $@"{xmlSaveAsPdfFolder}{settingsModel.XmlProjectFile}";
            xmlEmploeeysFile = $@"{xmlSaveAsPdfFolder}{settingsModel.XmlEmployeesFile}";

            DateTime date = DateTime.Now;

            //txtSaveLocation.Text = settingsModel.ProjectRootFolders.ToString().Replace($@"{settingsModel.ProjectRootTag}\", settingsModel.ProjectRootFolders.FullName);

            if (settingsModel.DefaultSavePath.ToString().Contains(settingsModel.DateTag))
            {
                txtSaveLocation.Text = txtSaveLocation.Text.Replace(settingsModel.DateTag, date.ToString("dd.MM.yyyy"));
            }


            if (settingsModel.ProjectRootFolders.Exists)
            {
                //TODO1: need to check it
                //Create .SaveAsPDF folder
                settingsModel.ProjectRootFolders.FullName.CreateHiddenFolder();
                if (File.Exists(xmlProjectFile))
                {
                    //load the XML file to projectModel model
                    projectModel = xmlProjectFile.XmlProjectFileToModel();

                    if (projectModel != null)
                    {
                        txtProjectName.Text = projectModel.ProjectName;
                        chkbSendNote.Checked = projectModel.NoteEmployee;
                        rtxtProjectNotes.Text = projectModel.ProjectNotes;
                    }
                }

                dgvEmployees.Rows.Clear();
                //load the XML file to Employees list-box
                if (File.Exists(xmlEmploeeysFile))
                {
                    employeesModel = xmlEmploeeysFile.XmlEmployeesFileToModel();
                    if (employeesModel != null)
                    {
                        foreach (EmployeeModel em in employeesModel)
                        {
                            dgvEmployees.Rows.Add(em.Id, em.FirstName, em.LastName, em.EmailAddress);
                        }
                    }
                }
            }

            tvFolders.Nodes.Clear();
            tvFolders.Nodes.Add(TreeHelper.CreateDirectoryNode(settingsModel.ProjectRootFolders));
            tvFolders.ExpandAll();
            tvFolders.SelectedNode = tvFolders.Nodes[0];
        }

        private void ClearForm()
        {
            txtProjectName.Clear();
            txtFullPath.Clear();
            txtSaveLocation.Clear();
            rtxtNotes.Clear();
        }
        private void btnOK_Click(object sender, EventArgs e)
        {
            if (ValidateForm())
            {
                if (!dataLoaded)
                {
                    LoadXmls();
                }

                #region Populate Models

                //build projectModel model

                projectModel.ProjectName = txtProjectName.Text;
                projectModel.ProjectNumber = txtProjectID.Text;
                projectModel.NoteEmployee = chkbSendNote.Checked;
                projectModel.ProjectNotes = rtxtProjectNotes.Text;

                //build the Employees model
                employeesModel = dgvEmployees.DgvEmployeesToModel();
                #endregion

                #region Create XML files for the models

                //create the SaveAsPDF hidden folder
                //xmlSaveAsPdfFolder.FullName.CreateHiddenFolder(); //already doing it on LoadXml() 

                //create projectModel XML file
                xmlProjectFile.ProjectModelToXmlFile(projectModel);

                //create the employeesModel XML file from List<EmployeeModel> 
                xmlEmploeeysFile.EmployeesModelToXmlFile(employeesModel);

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
                //save the mailItem to the current working directory and create an attachment list to add to PDF/mail   

                List<string> attList = new List<string>();
                //convert the message to HTML 
                if (mailItem.BodyFormat != OlBodyFormat.olFormatHTML)
                {
                    mailItem.BodyFormat = OlBodyFormat.olFormatHTML;
                }

                //Save the attachments and returning the actual file name list
                attList.AddRange(mailItem.SaveAttachments(dgvAttachments, txtSaveLocation.Text, false));

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
                mailItem.HTMLBody = tableStyle +
                                    projData +
                                    empString +
                                    attString +
                                    "</table></body>" +
                                    mailItem.HTMLBody;


                OfficeHelpers.SaveToPdf(mailItem, txtSaveLocation.Text);
            }

            Close();

            if (chbOpenPDF.Checked)
            {
                //TODO1:0 Open the PDF file
                MessageBox.Show("open PDF =" + chbOpenPDF.Checked.ToString());
            }

            settingsModel.OpenPDF = chbOpenPDF.Checked;
            //settingsModel.Save();
            //TODO: save settings 

        }
        private bool ValidateForm()
        {
            bool output = true;
            if (string.IsNullOrEmpty(txtProjectID.Text))
            {
                output = false;
            }
            if (string.IsNullOrEmpty(txtProjectName.Text))
            {
                output = false;
            }
            return output;
        }

        private void BtnSettings_Click(object sender, EventArgs e)
        {
            frmSettings frm = new frmSettings(this);
            frm.ShowDialog(this);
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
            foreach (EmployeeModel employee in employeesModel)
            {
                SendEmailToEmployee(employee.EmailAddress);
            }

        }

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
            settings = SettingsHelpers.loadSettingsToModel(settings);

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
                employeesModel.Add(model);
                dgvEmployees.Rows.Add(model.Id.ToString(),
                                        model.FirstName,
                                        model.LastName,
                                        model.EmailAddress);
            }

        }

        private void tvFolders_BeforeExpand(object sender, TreeViewCancelEventArgs e)
        {
            if (e.Node.Nodes.Count > 0)
            {
                if (e.Node.Nodes[0].Text == "..." && e.Node.Nodes[0].Tag == null)
                {
                    e.Node.Nodes.Clear();

                    //get the list of sub directory
                    string[] dirs = Directory.GetDirectories(e.Node.Tag.ToString());

                    foreach (string dir in dirs)
                    {
                        DirectoryInfo di = new DirectoryInfo(dir);
                        TreeNode node = new TreeNode(di.Name, 0, 1);

                        try
                        {
                            //keep the directory's full path in the tag for use later
                            node.Tag = dir;

                            //if the directory has sub directories add the place holder
                            if (di.GetDirectories().Count() > 0)
                                node.Nodes.Add(null, "...", 0, 0);
                        }
                        catch (UnauthorizedAccessException)
                        {
                            //display a locked folder icon
                            node.ImageIndex = 12;
                            node.SelectedImageIndex = 12;
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(ex.Message, "SaveAsPDF: treeView1_BeforeExpand",
                                MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                        finally
                        {
                            e.Node.Nodes.Add(node);
                        }
                    }
                }
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
            frmNewProject frm = new frmNewProject();
            frm.ShowDialog(this);
        }

        public void NewProjectComplete(ProjectModel model)
        {
            projectModel = model;

            txtProjectID.Text = projectModel.ProjectNumber;
            txtProjectName.Text = projectModel.ProjectName;
            rtxtProjectNotes.Text = projectModel.ProjectNotes;
            //TODO1: refresh folder treeview

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

            txtSaveLocation.Text = $@"{settingsModel.ProjectRootFolders.FullName.Trim('\\')}{settingsModel.ProjectRootFolders.ToString().Replace(
                                     settingsModel.ProjectRootTag, string.Empty)}";// no need the '\'
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

                    DirectoryInfo directoryInfo = new DirectoryInfo(Path.Combine(settingsModel.ProjectRootFolders.Parent.FullName, e.Node.FullPath));  // old path
                    directoryInfo.RnDir($@"{settingsModel.ProjectRootFolders.Parent.FullName}\{e.Node.Parent.FullPath}\{nodeNewLable}"); //nodeNewLable = new SAFE name

                    string specificPath = $@"{e.Node.Parent.FullPath}\{e.Label}";

                    TreeNode existingNode = TreeHelper.FindNodeByPath(tvFolders.Nodes, e.Node.FullPath);
                    if (existingNode != null)
                    {
                        TreeNode newNode = new TreeNode(nodeNewLable);
                        existingNode.Parent.Nodes.Insert(existingNode.Index, newNode.Text);
                        existingNode.Remove();

                    }
                }
                catch (Exception ex)
                {

                    MessageBox.Show(ex.Message + "\n" + Path.Combine(settingsModel.ProjectRootFolders.Parent.FullName, e.Node.FullPath), "SaveAsPDF:tvFolders_AfterLabelEdit");
                }

            }

        }



        private void chbOpenPDF_CheckedChanged(object sender, EventArgs e)
        {
            //TODO1: make sure it works 
            //settingsModel.OpenPDF = chbOpenPDF.Checked;
            Settings.Default.OpenPDF = chbOpenPDF.Checked;
            Settings.Default.Save();

        }

        private void tvFolders_NodeMouseDoubleClick(object sender, TreeNodeMouseClickEventArgs e)
        {
            Process.Start($@"{frmMain.settingsModel.ProjectRootFolders.Parent.FullName}\{e.Node.FullPath}");
        }

        private void tvFolders_NodeMouseClick(object sender, TreeNodeMouseClickEventArgs e)
        {
            //mySelectedNode = e.Node;
            //MessageBox.Show($"e.node.Name:{e.Node.Text} mySelectedNode: {mySelectedNode.FullPath}");
        }



        private void txtProjectID_Validating(object sender, System.ComponentModel.CancelEventArgs e)
        {
            if (!FileFoldersHelper.SafeProjectID(txtProjectID.Text))
            {
                e.Cancel = true;
                txtProjectID.Select(0, txtProjectID.Text.Length);
                txtProjectID.BackColor = System.Drawing.Color.Red;
                tsslStatus.Text = "מספר פרויקט לא חוקי";
                picBoxProjectID.Visible = true;
            }
        }

        private void txtProjectID_Validated(object sender, EventArgs e)
        {
            txtProjectID.BackColor = System.Drawing.Color.White;
            picBoxProjectID.Visible = false;
            tsslStatus.Text = string.Empty;

            ClearForm();
            LoadXmls();
            dataLoaded = true;
        }
    }
}