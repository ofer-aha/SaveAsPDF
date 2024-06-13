// Ignore Spelling: // Ignore Spelling: frm מפשחה הכל יש לבחור הודעות דואר אלקטרוני בלבד אימייל הסר הכול שם קובץ גודל יש לבחור הודעות דואר אלקטרוני בלבד ההודעה נשמרה ב  תאריך  שמירה  שם הפרויקט  מס פרויקט  הערות  שם משתמש בחר  הסר  מספר פרויקט כפי שמופיע במסטרפלן שם לא חוקי  אין להשתמש בתווים הבאים  עריכת שם שם לא חוקי לא ניתן ליצור שם ריק חובה תו אחד לפחות עריכת שם מספר פרויקט לא חוקי

using SaveAsPDF.Helpers;
using SaveAsPDF.Models;
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
        private readonly ISettingsRequester callingForm;
        private TreeNode mySelectedNode;
        private bool _isDirty = false;
        SettingsModel settingsModel = new SettingsModel();

        public frmSettings(ISettingsRequester caller)
        {
            InitializeComponent();
            callingForm = caller;


            //List<string> fList = new List<string>();
            //foreach (TreeNode node in tvProjectSubFolders.Nodes)
            //{
            //    fList.AddRange(TreeHelper.ListNodesPath(node));
            //}
            //settingsModel.ProjectRootFolders = fList;
        }

        private void frmSettings_Load(object sender, EventArgs e)
        {
            //TODO1:frmSettings_Load
            // ON HOLD

            tvProjectSubFolders.Nodes.Add("מספר_פרויקט");
            tvProjectSubFolders.HideSelection = false;
            tvProjectSubFolders.PathSeparator = @"\";

            txtRootFolder.Text = settingsModel.RootDrive;
            txtMinAttSize.Text = settingsModel.MinAttachmentSize.ToString();
            tvProjectSubFolders.LoadDefaultTree();
            tvProjectSubFolders.SelectedNode = tvProjectSubFolders.Nodes[0];

            cmbDefaultFolder.Items.Clear();

            TreeHelper.TvNodesToCombo(cmbDefaultFolder, tvProjectSubFolders.Nodes[0]);
            cmbDefaultFolder.SelectedIndex = settingsModel.DefaultFolderID;

            _isDirty = false;


        }


        //public frmSettings()
        //{
        //    InitializeComponent();

        //    tvProjectSubFolders.Nodes.Add("מספר_פרויקט");
        //    tvProjectSubFolders.HideSelection = false;
        //    tvProjectSubFolders.PathSeparator = @"\";

        //    txtRootFolder.Text = frmMain.settingsModel.RootDrive;
        //    txtMinAttSize.Text = frmMain.settingsModel.MinAttachmentSize.ToString();
        //    tvProjectSubFolders.LoadDefaultTree();
        //    tvProjectSubFolders.SelectedNode = tvProjectSubFolders.Nodes[0];

        //    //List<string> fList = new List<string>();

        //    //foreach (TreeNode node in tvProjectSubFolders.Nodes)
        //    //{
        //    //    fList.AddRange(TreeHelper.ListNodesPath(node));
        //    //}
        //    //settingsModel.ProjectRootFolders = fList;

        //    cmbDefaultFolder.Items.Clear();

        //    TreeHelper.TvNodesToCombo(cmbDefaultFolder, tvProjectSubFolders.Nodes[0]);
        //    cmbDefaultFolder.SelectedIndex = frmMain.settingsModel.DefaultFolderID;

        //    _isDirty = false;
        //}

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
            Close();
        }
        /// <summary>
        /// save the settings form the model to settings.settings 
        /// </summary>
        private void SaveSettings()
        {
            _isDirty = false;
            SettingsModel settingsModel = new SettingsModel();
            settingsModel.RootDrive = txtRootFolder.Text;
            settingsModel.MinAttachmentSize = int.Parse(txtMinAttSize.Text);
            settingsModel.DefaultSavePath = cmbDefaultFolder.Text;

            //save the new settings
            SaveDefaultTree();
            //settingsModel.Save();


            SettingsHelpers.saveModelToSettings();

            callingForm.SettingsComplete(settingsModel);
        }

        private void btnSaveSettings_Click(object sender, EventArgs e)
        {
            SaveSettings();
        }


        private void btnFolders_Click(object sender, EventArgs e)
        {
            var Dialog = new FolderPicker();

            Dialog.InputPath = frmMain.settingsModel.RootDrive;

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

            string defaultTreeFileName = $@"{directoryPath}\{settingsModel.DefaultTreeFile}";

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
        //    settingsModel settingsModel = new settingsModel();
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

            tvProjectSubFolders.AddNode(mySelectedNode, frmMain.settingsModel.DateTag);
            _isDirty = true;

        }

        private void menuAppendDate_Click(object sender, EventArgs e)
        {
            //tvProjectSubFolders.SelectedNode.Name += settingsModel.dateTag;
            TreeHelper.RenameNode(tvProjectSubFolders, mySelectedNode, mySelectedNode.Name + frmMain.settingsModel.DateTag);
        }

        private void cmbDefaultFolder_SelectedIndexChanged(object sender, EventArgs e)
        {
            //update settings 
            frmMain.settingsModel.DefaultFolderID = (int)cmbDefaultFolder.SelectedIndex;
            frmMain.settingsModel.ProjectRootFolders = new DirectoryInfo(cmbDefaultFolder.Text);
            _isDirty = true;
        }

        private void txtRootFolder_TextChanged(object sender, EventArgs e)
        {
            _isDirty = true;
        }
    }
}