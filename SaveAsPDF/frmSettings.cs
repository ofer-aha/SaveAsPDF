// Ignore Spelling: frm

using SaveAsPDF.Helpers;
using SaveAsPDF.Models;
using SaveAsPDF.Properties;
using System;
using System.IO;
using System.Reflection;
using System.Windows.Controls;
using System.Windows.Forms;
using System.Xml;


namespace SaveAsPDF
{
    public partial class frmSettings : Form
    {
        private TreeNode mySelectedNode;
        private bool _isDirty = false;

        private void frmSettings_Load(object sender, EventArgs e)
        {
            //TODO1:frmSettings_Load
            // ON HOLD


        }


        public frmSettings()
        {
            InitializeComponent();

            tvProjectSubFolders.Nodes.Add("מספר_פרויקט");
            tvProjectSubFolders.HideSelection = false;
            tvProjectSubFolders.PathSeparator = @"\";

            txtRootFolder.Text = Settings.Default.RootDrive;
            txtMinAttSize.Text = Settings.Default.MinAttachmentSize.ToString();
            tvProjectSubFolders.LoadDefaultTree();
            tvProjectSubFolders.SelectedNode = tvProjectSubFolders.Nodes[0];

            //TODO: reconsider the need for SettingsModel
            //SettingsModel settingsModel = new SettingsModel();
            //List<string> fList = new List<string>();

            //foreach (TreeNode node in tvProjectSubFolders.Nodes)
            //{
            //    fList.AddRange(TreeHelper.ListNodesPath(node));
            //}
            //settingsModel.ProjectRootFolders = fList;

            cmbDefaultFolder.Items.Clear();

            TreeHelper.TvNodesToCombo(cmbDefaultFolder, tvProjectSubFolders.Nodes[0]);
            cmbDefaultFolder.SelectedIndex = Settings.Default.DefaultFolderID;
            _isDirty = false;
        }



        private void bntCancel_Click(object sender, EventArgs e)
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
                _isDirty = true;
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
                    MessageBox.Show("שם לא חוקי.\n לא ניתן ליצור שם ריק. חובה תו אחד לפחות", "עריכת שם");
                    e.Node.BeginEdit();
                }
            }
        }

        private void menueAdd_Click(object sender, EventArgs e)
        {
            tvProjectSubFolders.AddNode(mySelectedNode);
            _isDirty = true;
        }

        private void menuDel_Click(object sender, EventArgs e)
        {
            tvProjectSubFolders.DelNode(mySelectedNode);
            _isDirty = true;
        }

        private void menuRename_Click(object sender, EventArgs e)
        {
            tvProjectSubFolders.RenameNode(mySelectedNode);
            _isDirty = true;
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            if (_isDirty)
            {
                DialogResult result = MessageBox.Show("שמור שינויים?", "SaveAsPDF", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (result == DialogResult.Yes)
                {
                    SaveSettings();
                }
            }
            _isDirty = false;
            Close();
        }

        private void SaveSettings()
        {
            //TODO1: Save settings is wrong!!!!
            _isDirty = false;

            Settings.Default.RootDrive = txtRootFolder.Text;
            Settings.Default.MinAttachmentSize = int.Parse(txtMinAttSize.Text);

            //save the new settings
            SaveDefaultTree();
            Settings.Default.Save();
        }

        private void btnSaveSettings_Click(object sender, EventArgs e)
        {
            SaveSettings();
        }


        private void btnFolders_Click(object sender, EventArgs e)
        {

            var Dialog = new FolderPicker();

            Dialog.InputPath = Settings.Default.RootDrive;

            if (Dialog.ShowDialog(Handle) == true)
            {
                _isDirty = true;
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

            string defaultTreeFileName = $@"{directoryPath}\{Settings.Default.DefaultTreeFile}";

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
            _isDirty = true;
        }

        private void btnSaveTreeFile_Click(object sender, EventArgs e)
        {
            SaveFileDialog dlg = new SaveFileDialog();
            dlg.Title = "שמור קובץ תיקיות";
            dlg.Filter = "קובץ תיקיות (*.fld)|*.fld";

            if (dlg.ShowDialog() == DialogResult.OK)
            {
                TreeHelper.SaveTreeViewIntoFile(dlg.FileName, tvProjectSubFolders);
                //TODO3: NEXT VERSION: change the tree file to XML
            }
            _isDirty = false;
        }

        private XmlNode GetXmlNode(TreeViewItem tnode, XmlDocument d)
        {
            XmlNode n = d.CreateNode(XmlNodeType.Element, tnode.Name, "");
            foreach (TreeViewItem t in tnode.Items)
            {
                n.AppendChild(GetXmlNode(t, d));
            }
            return n;
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
        //    settingsModel.ProjectRootFolders = fList;

        //    cbDefaultFolder.DataSource = settingsModel.ProjectRootFolders;
        //    cbDefaultFolder.Refresh();

        //}


        private void btnLoadDefaultTree_Click(object sender, EventArgs e)
        {
            tvProjectSubFolders.LoadDefaultTree();
            _isDirty = true;
        }


        private void txtMinAttSize_TextChanged(object sender, EventArgs e)
        {
            _isDirty = true;
        }

        private void tvProjectSubFolders_AfterSelect(object sender, TreeViewEventArgs e)
        {
        }


        private void menuAddDate_Click(object sender, EventArgs e)
        {

            tvProjectSubFolders.AddNode(mySelectedNode, Settings.Default.DateTag);
            _isDirty = true;

        }

        private void menuAppendDate_Click(object sender, EventArgs e)
        {
            //tvProjectSubFolders.SelectedNode.Name += Settings.Default.dateTag;
            TreeHelper.RenameNode(tvProjectSubFolders, mySelectedNode, mySelectedNode.Name + Settings.Default.DateTag);
        }

        private void cmbDefaultFolder_SelectedIndexChanged(object sender, EventArgs e)
        {
            //update settings 
            Settings.Default.DefaultFolderID = cmbDefaultFolder.SelectedIndex;
            Settings.Default.DefaultFolderSettings = cmbDefaultFolder.Text;
            _isDirty = true;
        }

        private void txtRootFolder_TextChanged(object sender, EventArgs e)
        {
            _isDirty = true;
        }
    }
}