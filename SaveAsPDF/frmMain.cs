using Microsoft.Office.Interop.Outlook;
using Microsoft.Office.Interop.Word;
using SaveAsPDF.Models;
using SaveAsPDF.Helpers;

using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Font = Microsoft.Office.Interop.Word.Font;
using Outlook = Microsoft.Office.Interop.Outlook;
using word = Microsoft.Office.Interop.Word;
using System.Configuration;

namespace SaveAsPDF
{
    public partial class frmMain : Form, IEmployeeRequester
    {
        private List<EmployeeModel> employees = new List<EmployeeModel>();
        private ProjectModel project = new ProjectModel();

        // construct the full path for evrithig
        private string sPath;
        private string xmlSaveAsPdfFolder;
        private string xmlProjectFile;
        private string xmlEmploeeysFile;

        public frmMain()
        {
            InitializeComponent();

           
            dgvEmployees.Columns[0].Visible = false;
            dgvEmployees.Columns[1].HeaderText = "שם פרטי";
            dgvEmployees.Columns[2].HeaderText = "שם מפשחה";
            dgvEmployees.Columns[3].HeaderText = "אימייל";
            dgvEmployees.Columns[4].Visible = false;


        }
        private void frmMain_Load(object sender, EventArgs e)
        {
            MailItem mi = null;
            MailItem mailItem = ThisAddIn.TypeOfMailitem(mi);


            if (mailItem is MailItem)
            {
                chkbSelectAllAttachments.Checked = true;
                chkbSelectAllAttachments.Text = "הסר הכל";
                txtSubject.Text = mailItem.Subject;
                txtSubject.ReadOnly = true;
                txtSubject.TextAlign = HorizontalAlignment.Right;

                //load the list from model
                dgvAttachments.DataSource = mailItem.AttachmetsToModel();

                dgvAttachments.Columns[0].Visible = false;
                dgvAttachments.Columns[1].HeaderText = "";
                dgvAttachments.Columns[2].HeaderText = "שם קובץ";
                dgvAttachments.Columns[2].ReadOnly = true;
                dgvAttachments.Columns[3].HeaderText = "גודל";
                dgvAttachments.Columns[3].ReadOnly = true;
            }
            else
            {
                MessageBox.Show("Not a Mail Item");
                this.Close();
            }
            //string tmpFoder = @System.IO.Path.GetTempPath();



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

                treeView1.Nodes.Add(node);
            }
        }


        private void btnCancel_Click(object sender, EventArgs e)
        {
            //close the form - do  nothing
            this.Close();
        }
        private void btnFolders_Click(object sender, EventArgs e)
        {
            

            var Dialog = new FolderPicker();

            if (!Directory.Exists(sPath))
            {
                Dialog.InputPath = GlobalConfig.rootDrive;
            }
            else
            {
                Dialog.InputPath = sPath;
            }

            if (Dialog.ShowDialog(Handle) == true)
            {
                
                treeView1.Nodes.Clear();
                treeView1.Nodes.Add(TreeHelper.TraverseDirectory(Dialog.ResultPath));
            }
        }
        private void txtProjectID_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                // construct the full path for evrithig
                sPath = txtProjectID.Text.Trim().ProjectFullPath();
                xmlSaveAsPdfFolder = $"{sPath}{GlobalConfig.xmlSaveAsPdfFolder}";
                xmlProjectFile = $"{xmlSaveAsPdfFolder}{GlobalConfig.xmlProjectFile}";
                xmlEmploeeysFile = $"{xmlSaveAsPdfFolder}{GlobalConfig.xmlEmploeeysFile}";



                //if project's folder don't exist - create it 
                //same for SaveAsPDF's default hidden subfolser 
                if (!Directory.Exists(xmlSaveAsPdfFolder))
                {
                    if (!Directory.Exists(sPath))
                    {
                        Directory.CreateDirectory(sPath);
                    }
                    FileFoldersHelper.CreateHiddenFolder(xmlSaveAsPdfFolder);
                }
                //load the XML file to project model
                if (File.Exists(xmlProjectFile))
                {
                    project = xmlProjectFile.XmlProjectFileToModel();
                    txtProjectName.Text = project.ProjectName;
                    chkbSendNote.Checked = project.NoteEmployee;
                }
                else
                {
                    project.ProjectNumber = txtProjectID.Text.Trim();
                    //TODO: form new project with default folders creation

                }

                dgvEmployees.Rows.Clear();
                //load the XML file to Emploeeys listbox
                if (File.Exists(xmlEmploeeysFile))
                {
                    employees = xmlEmploeeysFile.XmlEmloyeesFileToModel();
                    
                    foreach (EmployeeModel em in employees)
                    {
                        dgvEmployees.Rows.Add(em.Id, em.FirstName, em.LastName, em.EmailAddress);
                    }
                }
                
                
                txtFullPath.Text = sPath;
                treeView1.Nodes.Clear();
                DirectoryInfo d = new DirectoryInfo(sPath);
                //treeView1.Nodes.Add(TreeHelper.TraverseDirectory(sPath));
                //treeView1.Nodes.AddRange(TreeHelper.getFolderNodes(sPath, true).ToArray());
                treeView1.Nodes.Add(TreeHelper.CreateDirectoryNode(d));

            }
        }

        private void btnOK_Click(object sender, EventArgs e)
        {

            //TODO: before doing ANYTHING make sure there is a project number to begin with

            #region MailItem

            MailItem mi = null;
            MailItem mailItem = ThisAddIn.TypeOfMailitem(mi);

            #endregion

            #region Populate Models

            //build project modele
            if (txtProjectName.Text != "")
            {
                project.ProjectName = txtProjectName.Text;
                project.ProjectNumber = txtProjectID.Text;
                project.NoteEmployee = chkbSendNote.Checked;
            }


            //build the Employees model
            //List<EmployeeModel> employees = new List<EmployeeModel>();
            employees = dgvEmployees.DgvEmployessToModel();
            #endregion

            #region Creat XML files for the models
                        
            //create the SaveAsPDF hidden folder
            FileFoldersHelper.CreateHiddenFolder(xmlSaveAsPdfFolder);

            //create project XML file
            XmlFileHelper.ProjectModelToXmlFile(project, xmlProjectFile);

            //create the employees XML file from List<EmployeeModel> 
            XmlFileHelper.EmployeesModelToXmlFile(employees, xmlEmploeeysFile);
            #endregion


            //TODO: inject more data to the file befor converting to PDF

            //save the mailItem to the current working directory. 
            //OfficeHelpers.SaveToPDF(mailItem, lblFolder.Text); 

            //TODO: what about attuchment saveing? 

        }


        private void BtnSettings_Click(object sender, EventArgs e)
        {
            frmSettings frm = new frmSettings();
            frm.ShowDialog(this);
        }

        private void dgvAttachments_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            //TODO: make somthing meaningfull
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
            SendEmailToEmployee("somthing");
        }

        private void SendEmailToEmployee(string text)
        {
           //TODO: send email
            MessageBox.Show($"Send email to {text}");
        }

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
            employees.Add(model);

            dgvEmployees.Rows.Add(model.Id.ToString(),
                                    model.FirstName,
                                    model.LastName,
                                    model.EmailAddress);
        }

        private void treeView1_BeforeExpand(object sender, TreeViewCancelEventArgs e)
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
                        catch (System.Exception ex)
                        {
                            MessageBox.Show(ex.Message, "DirectoryLister",
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

    }
}
