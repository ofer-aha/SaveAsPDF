using SaveAsPDF.Helpers;
using SaveAsPDF.Models;
using SaveAsPDF.Properties;
using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Windows.Forms;


namespace SaveAsPDF
{
    public partial class frmSettings : Form
    {
        private TreeNode mySelectedNode;
        private bool formChanged = false; 
        public frmSettings()
        {
            InitializeComponent();
            tvProjectSubFolders.Nodes.Add("מספר_פרויקט");
            tvProjectSubFolders.HideSelection = false;
            tvProjectSubFolders.PathSeparator = @"\";
            LoadSettings();
            
        }

        private void LoadSettings()
        {

            txtRootFolder.Text = Settings.Default.rootDrive; 
            txtDefaultFolfer.Text = Settings.Default.defaultFolder;
            txtMinAttSize.Text = Settings.Default.minAttachmentSize.ToString();
            TreeHelper.LoadDefaultTree(tvProjectSubFolders);

            SettingsModel settingsModel = new SettingsModel();
            List<string> fList = new List<string>();

            foreach (TreeNode node in tvProjectSubFolders.Nodes)
            {
                fList.AddRange(TreeHelper.ListNodesPath(node));
            }
            settingsModel.ProjectFolders = fList;

            cbDefaultFolder.DataSource = settingsModel.ProjectFolders;
            cbDefaultFolder.Refresh();
            formChanged = false;

        }



        private void bntCancel_Click(object sender, System.EventArgs e)
        {
            this.Close();
        }

        /// <summary>
        /// Get the tree node under the mouse pointer and 
        /// save it in the mySelectedNode variable.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void tvProjectSubFolders_MouseDown(object sender, MouseEventArgs e)
        {
            mySelectedNode = tvProjectSubFolders.GetNodeAt(e.X, e.Y);
        }

        private void tvProjectSubFolders_AfterLabelEdit(object sender, NodeLabelEditEventArgs e)
        {
            if (e.Label != null)
            {
                formChanged = true;
                if (e.Label.Length > 0)
                {

                    if (e.Label.IndexOfAny(new char[] { '\\', '/', ':', '*', '?', '<', '>', '|', '"' }) == -1)
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
                           "אין להשתמש בתווים הבאים '\\', '/', ':', '*', '?', '<', '>', '|', '\"' ",
                           "עריכת שם");
                        e.Node.BeginEdit();
                    }
                }
                else
                {
                    // Cancel the label edit action, inform the user, and
                    // place the node in edit mode again. 
                    e.CancelEdit = true;
                    MessageBox.Show("שם לא חוקי.\nלא ניתן ליצור שם ריק. חובה תו אחד לפחות",
                       "עריכת שם");
                    e.Node.BeginEdit();
                }
            }
        }

        private void menueAdd_Click(object sender, System.EventArgs e)
        {
            TreeHelper.AddNode(tvProjectSubFolders, mySelectedNode);
            formChanged = true;    
        }

        private void menuDel_Click(object sender, System.EventArgs e)
        {
            TreeHelper.DelNode(tvProjectSubFolders, mySelectedNode);
            formChanged = true;    
        }

        private void menuRename_Click(object sender, System.EventArgs e)
        {
            TreeHelper.RenameNode(tvProjectSubFolders, mySelectedNode);
            formChanged = true;
        }

        private void btnOK_Click(object sender, System.EventArgs e)
        {
            if (formChanged)
            {
                
                DialogResult result = MessageBox.Show("שמור שינויים?","SaveAsPDF",MessageBoxButtons.YesNo,MessageBoxIcon.Question);
                if (result == DialogResult.Yes )
                {
                    SaveSettings();
                }

            }
    
            Close();

        }

        private void SaveSettings()
        {
           
            formChanged = false;

            Settings.Default.defaultFolder = txtDefaultFolfer.Text;
            Settings.Default.rootDrive = txtRootFolder.Text;
            Settings.Default.minAttachmentSize = int.Parse(txtMinAttSize.Text);

            //save the new settings
            
            Settings.Default.Save();
            Settings.Default.Upgrade();

        }

        private void btnDefaultTree_Click(object sender, System.EventArgs e)
        {
            TreeHelper.RestTree(tvProjectSubFolders);
            SettingsModel settingsModel = new SettingsModel();
            List<string> fList = new List<string>();

            foreach (TreeNode node in tvProjectSubFolders.Nodes)
            {
                fList.AddRange(TreeHelper.ListNodesPath(node));
            }
            settingsModel.ProjectFolders = fList;

            cbDefaultFolder.DataSource = settingsModel.ProjectFolders;
            cbDefaultFolder.Refresh();

        }


        private void btnSaveSettings_Click(object sender, System.EventArgs e)
        {
            SaveSettings();
        }


        private void btnFolders_Click(object sender, EventArgs e)
        {

            var Dialog = new FolderPicker();
            
                Dialog.InputPath = Settings.Default.rootDrive;

            if (Dialog.ShowDialog(Handle) == true)
            {
                formChanged = true;
                txtRootFolder.Text = Dialog.ResultPath;

            }
        }

        private void btnSaveDefaultTree_Click(object sender, EventArgs e)
        {
            Assembly assembly = Assembly.GetExecutingAssembly();
            string location = assembly.CodeBase;
            string fullPath = new Uri(location).LocalPath; // path including the dll 
            string directoryPath = Path.GetDirectoryName(fullPath); // directory path 

            string defaultTreeFileName = directoryPath + "\\" + Settings.Default.defaultTreeFile;

            TreeHelper.SaveTreeViewIntoFile(defaultTreeFileName, tvProjectSubFolders);


        }

        private void btnLoadTreeFile_Click_1(object sender, EventArgs e)
        {
            OpenFileDialog dlg = new OpenFileDialog();
            dlg.Title = "פתח קובץ תיקיות";
            dlg.Filter = "קובץ תיקיות (*.fld)|*.fld";

            if (dlg.ShowDialog() == DialogResult.OK)
            {
                TreeHelper.LoadTreeViewFromFile(dlg.FileName, tvProjectSubFolders);
            }

        }

        private void btnSaveTreeFile_Click(object sender, EventArgs e)
        {
            SaveFileDialog dlg = new SaveFileDialog();
            dlg.Title = "שמור קובץ תיקיות";
            dlg.Filter = "קובץ תיקיות (*.fld)|*.fld";

            if (dlg.ShowDialog() == DialogResult.OK)
            {
                TreeHelper.SaveTreeViewIntoFile(dlg.FileName, tvProjectSubFolders);
            }
        }

        private void btnRestTree_Click(object sender, EventArgs e)
        {
            tvProjectSubFolders.Nodes.Clear();
            TreeHelper.RestTree(tvProjectSubFolders);
        }


        private void btnLoadDefaultTree_Click(object sender, EventArgs e)
        {
            TreeHelper.LoadDefaultTree(tvProjectSubFolders);
        }

        private void cbDefaultFolder_SelectionChangeCommitted(object sender, EventArgs e)
        {
            txtDefaultFolfer.Text = cbDefaultFolder.SelectedItem.ToString();

            formChanged = true;

        }

        private void txtMinAttSize_TextChanged(object sender, EventArgs e)
        {
            formChanged = true;
        }
    }
}