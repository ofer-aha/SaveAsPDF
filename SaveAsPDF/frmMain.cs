// Ignore Spelling: frm

using Microsoft.Office.Interop.Outlook;
using SaveAsPDF.Helpers;
using SaveAsPDF.Models;
using SaveAsPDF.Properties;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;
using Exception = System.Exception;

namespace SaveAsPDF
{

    public partial class frmMain : Form, IEmployeeRequester, INewProjectRequester
    {
        private List<EmployeeModel> employees = new List<EmployeeModel>();
        private ProjectModel project = new ProjectModel();

        // construct the full path for evrithig
        public static DirectoryInfo sPath;
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

            dgvEmployees.Columns[0].Visible = false;
            dgvEmployees.Columns[1].HeaderText = "שם פרטי";
            dgvEmployees.Columns[2].HeaderText = "שם מפשחה";
            dgvEmployees.Columns[3].HeaderText = "אימייל";
            //dgvEmployees.Columns[4].Visible = false;

            //Load the context menue to the rich textboxes 
            rtxtNotes.EnableContextMenu();
            rtxtProjectNotes.EnableContextMenu();
            txtFullPath.EnableContextMenu();
            txtProjectID.EnableContextMenu();
            txtProjectName.EnableContextMenu();
            tvFolders.EnableContextMenu();


        }
        private void frmMain_Load(object sender, EventArgs e)
        {

            sPath = txtProjectID.Text.ProjectFullPath();

            if (mailItem is MailItem && mailItem != null)
            {
                chkbSelectAllAttachments.Checked = true;
                chkbSelectAllAttachments.Text = "הסר הכל";
                txtSubject.Text = mailItem.Subject;
                txtSaveLocation.Text = sPath.FullName;

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

            if (!sPath.Exists)
            {
                Dialog.InputPath = Settings.Default.rootDrive;
            }
            else
            {
                Dialog.InputPath = sPath.FullName;
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


            }
        }
        /// <summary>
        /// Load the XML files for this project to the modules
        /// </summary>
        private void LoadXmls()
        {
            // construct the full path for evrything
            sPath = txtProjectID.Text.Trim().ProjectFullPath();
            xmlSaveAsPdfFolder = new DirectoryInfo(System.IO.Path.Combine(sPath.FullName, Settings.Default.xmlSaveAsPdfFolder));
            xmlProjectFile = $"{xmlSaveAsPdfFolder}{Settings.Default.xmlProjectFile}";
            xmlEmploeeysFile = $"{xmlSaveAsPdfFolder}{Settings.Default.xmlEmploeeysFile}";


            DateTime date = DateTime.Now;

            txtSaveLocation.Text = Settings.Default.defaultFolder.Replace($"{Settings.Default.projectRootTag}\\", sPath.FullName);

            if (Settings.Default.defaultFolder.Contains(Settings.Default.dateTag))
            {
                txtSaveLocation.Text = txtSaveLocation.Text.Replace(Settings.Default.dateTag, date.ToString("dd.MM.yyyy"));
            }


            if (sPath.Exists)
            {
                //Create .SaveAsPDF
                sPath.FullName.CreateHiddenFolder();
                if (File.Exists(xmlProjectFile))
                {
                    //load the XML file to project model
                    project = xmlProjectFile.XmlProjectFileToModel();

                    if (project != null)
                    {
                        txtProjectName.Text = project.ProjectName;
                        chkbSendNote.Checked = project.NoteEmployee;
                        rtxtProjectNotes.Text = project.ProjectNotes;
                    }
                }
                dgvEmployees.Rows.Clear();
                //load the XML file to Emploeeys listbox
                if (File.Exists(xmlEmploeeysFile))
                {
                    employees = xmlEmploeeysFile.XmlEmloyeesFileToModel();
                    if (employees != null)
                    {
                        foreach (EmployeeModel em in employees)
                        {
                            dgvEmployees.Rows.Add(em.Id, em.FirstName, em.LastName, em.EmailAddress);
                        }
                    }
                }
            }

            txtFullPath.Text = sPath.FullName;
            tvFolders.Nodes.Clear();
            tvFolders.Nodes.Add(TreeHelper.CreateDirectoryNode(sPath));
            tvFolders.ExpandAll();
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

                //build project modele

                project.ProjectName = txtProjectName.Text;
                project.ProjectNumber = txtProjectID.Text;
                project.NoteEmployee = chkbSendNote.Checked;
                project.ProjectNotes = rtxtProjectNotes.Text;

                //build the Employees model
                employees = dgvEmployees.DgvEmployeesToModel();
                #endregion

                #region Creat XML files for the models

                //create the SaveAsPDF hidden folder
                xmlSaveAsPdfFolder.FullName.CreateHiddenFolder();

                //create project XML file
                xmlProjectFile.ProjectModelToXmlFile(project);

                //create the employees XML file from List<EmployeeModel> 
                xmlEmploeeysFile.EmployeesModelToXmlFile(employees);

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
                    Dialog.InputPath = Settings.Default.rootDrive;
                    if (Dialog.ShowDialog(Handle) == true)
                    {
                        txtSaveLocation.Text = Dialog.ResultPath;
                    }
                }
                //save the mailItem to the current working directory and create an attachment list to add to pdf/mail   

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

            Settings.Default.OpenPDF = chbOpenPDF.Checked;
            Settings.Default.Save();

        }
        /// <summary>
        /// making sure nothing is missing before closeing the form
        /// and creating the PDF file
        /// </summary>
        /// <returns>False if validation fails</returns>
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
            frmSettings frm = new frmSettings();
            frm.ShowDialog(this);
        }

        private void dgvAttachments_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            //TODO1: Open attachment when double clicking it on the list
            //1. save attachment to temp folder 
            //string tmpFoder = @System.IO.Path.GetTempPath();
            //2. exec. the file using default file asociating 
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
            foreach (EmployeeModel employee in employees)
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
                employees.Add(model);
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

                    //get the list of sub direcotires
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
            project = model;

            txtProjectID.Text = project.ProjectNumber;
            txtProjectName.Text = project.ProjectName;
            rtxtProjectNotes.Text = project.ProjectNotes;
            //TODO1: refresh folder treeview

        }

        private void inputBox_Validating(object sender, InputBoxValidatingArgs e)
        {
            if (e.Text.Trim().Length == 0)
            {
                e.Cancel = true;
                e.Message = "חובה למלא";
            }
        }

        private void txtProjectID_MouseHover(object sender, EventArgs e)
        {
            tsslStatus.Enabled = true;
            tsslStatus.Text = "מספר פרויקט כפי שמופיע במסטרפלן";
        }

        private void btnCopyNotesToMail_Click(object sender, EventArgs e)
        {
            rtxtNotes.Text = rtxtNotes.Text + "\n" + rtxtProjectNotes.Text;
        }

        private void btnCopyNotesToProject_Click(object sender, EventArgs e)
        {
            rtxtProjectNotes.Text = rtxtProjectNotes.Text + "\n" + rtxtNotes.Text;
        }

        private void tvFolders_AfterSelect(object sender, TreeViewEventArgs e)
        {
            TreeNode CurrentNode = e.Node;
            string fullpath = CurrentNode.FullPath;
            mySelectedNode = CurrentNode;

            txtSaveLocation.Text = sPath.Parent.FullName + "\\" + fullpath;


        }

        private void tvFolders_MouseDown(object sender, MouseEventArgs e)
        {
            //mySelectedNode = tvFolders.GetNodeAt(e.X, e.Y);
            //txtSaveLocation.Text = Settings.Default.rootDrive + mySelectedNode.FullPath;
        }

        private void tvFolders_AfterLabelEdit(object sender, NodeLabelEditEventArgs e)
        {
            //TODO1:_0 Make sure user entered safe folder name 

            string inTXT = e.Label.SafeFileName(); ;

            if (inTXT != null)
            {
                if (inTXT.Length > 0)
                {


                    if (inTXT.IndexOfAny(new char[] { '\\', '/', ':', '*', '?', '<', '>', '|', '"' }) == -1)

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
                    MessageBox.Show("שם לא חוקי.\nלא ניתן ליצור שם ריק. חובה תו אחד לפחות", "עריכת שם");
                    e.Node.BeginEdit();
                }

                try
                {
                    DirectoryInfo directoryInfo = new DirectoryInfo(Path.Combine(sPath.Parent.FullName, e.Node.FullPath));
                    directoryInfo.RnDir($"{sPath.Parent.FullName}\\{mySelectedNode.Parent.FullPath}\\{inTXT}"); //inTXT = new name
                }
                catch (Exception ex)
                {

                    MessageBox.Show(ex.Message + "\n" + System.IO.Path.Combine(sPath.Parent.FullName, mySelectedNode.FullPath), "SaveAsPDF:tvFolders_AfterLabelEdit");
                }

            }

        }



        private void chbOpenPDF_CheckedChanged(object sender, EventArgs e)
        {
            //TODO1: make sure it works 
            Settings.Default.OpenPDF = chbOpenPDF.Checked;

        }
    }
}
