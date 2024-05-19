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
            txtDefaultFolder.Text = Settings.Default.defaultFolder;
            txtMinAttSize.Text = Settings.Default.minAttachmentSize.ToString();
            tvProjectSubFolders.LoadDefaultTree();

            SettingsModel settingsModel = new SettingsModel();
            List<string> fList = new List<string>();

            foreach (TreeNode node in tvProjectSubFolders.Nodes)
            {
                fList.AddRange(TreeHelper.ListNodesPath(node));
            }
            settingsModel.ProjectFolders = fList;
            formChanged = false;

        }



        private void bntCancel_Click(object sender, System.EventArgs e)
        {
            Close();
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
            tvProjectSubFolders.AddNode(mySelectedNode);
            formChanged = true;    
        }

        private void menuDel_Click(object sender, System.EventArgs e)
        {
            tvProjectSubFolders.DelNode(mySelectedNode);
            formChanged = true;    
        }

        private void menuRename_Click(object sender, System.EventArgs e)
        {
            tvProjectSubFolders.RenameNode(mySelectedNode);
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
            //TODO1: Save settings is worng!!!!
            formChanged = false;

            Settings.Default.defaultFolder = txtDefaultFolder.Text;
            Settings.Default.rootDrive = txtRootFolder.Text;
            Settings.Default.minAttachmentSize = int.Parse(txtMinAttSize.Text);

            //save the new settings
            SaveDefaultTree();
            Properties.Settings.Default.Save();
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
            SaveDefaultTree();

        }

        private void SaveDefaultTree()
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
                tvProjectSubFolders.LoadTreeViewFromFile(dlg.FileName);
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

        //TODO3: to check
        //private void btnDefaultTree_Click(object sender, System.EventArgs e)
        //{
        //    TreeHelper.RestTree(tvProjectSubFolders);
        //    SettingsModel settingsModel = new SettingsModel();
        //    List<string> fList = new List<string>();

        //    foreach (TreeNode node in tvProjectSubFolders.Nodes)
        //    {
        //        fList.AddRange(TreeHelper.ListNodesPath(node));
        //    }
        //    settingsModel.ProjectFolders = fList;

        //    cbDefaultFolder.DataSource = settingsModel.ProjectFolders;
        //    cbDefaultFolder.Refresh();

        //}


        private void btnLoadDefaultTree_Click(object sender, EventArgs e)
        {
            tvProjectSubFolders.LoadDefaultTree();
        }


        private void txtMinAttSize_TextChanged(object sender, EventArgs e)
        {
            formChanged = true;
        }

        private void tvProjectSubFolders_AfterSelect(object sender, TreeViewEventArgs e)
        {
        }

        private void btnFolderSelect_Click(object sender, EventArgs e)
        {
            //TODO1: 1 fix the folder reading when its empty 
            if (!string.IsNullOrEmpty(txtDefaultFolder.Text))
            {

            TreeNode CurrentNode = tvProjectSubFolders.SelectedNode;
            string fullpath = CurrentNode.FullPath;
            txtDefaultFolder.Text = fullpath;

            formChanged = true;
            }
            else 
            {
                txtDefaultFolder.Text = "_מספר_פרויקט_";
            }

        }

        private void menuAddDate_Click(object sender, EventArgs e)
        {
            
            tvProjectSubFolders.AddNode(mySelectedNode, Settings.Default.dateTag);  
            formChanged = true;

        }

        private void menuAppendDate_Click(object sender, EventArgs e)
        {
            //tvProjectSubFolders.SelectedNode.Name += Settings.Default.dateTag;
            TreeHelper.RenameNode(tvProjectSubFolders, mySelectedNode, mySelectedNode.Name + Settings.Default.dateTag);
        }
    }
}