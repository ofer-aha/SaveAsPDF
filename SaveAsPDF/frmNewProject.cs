using SaveAsPDF.Helpers;
using SaveAsPDF.Models;
using SaveAsPDF.Properties;
using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;


namespace SaveAsPDF
{
    public partial class frmNewProject : Form
    {
        private readonly INewProjectRequester callingForm;
        private List<string> subFolfers = new List<string>();
        private TreeNode mySelectedNode;
        public frmNewProject()
        {
            InitializeComponent();

            LoadSettings();
        }

        private void LoadSettings()
        {
            TreeHelper.LoadDefaultTree(tvDefaultSubFolders);

        }

        private void btmNewProject_Click(object sender, EventArgs e)
        {
            //populate ProjectModel
            if (ValidateForm())
            {
                ProjectModel p = new ProjectModel
                {
                    ProjectName = txtProjectName.Text,
                    ProjectNumber = txtProjectId.Text.Trim(),
                    ProjectNotes = txtProjectNotes.Text
                };

                //create the project's subfolders 

                subFolfers = TreeHelper.ListNodesPath(tvDefaultSubFolders.Nodes[0]);
                foreach (string subFolder in subFolfers)
                {
                    FileFoldersHelper.MkDir(subFolder);
                }

                callingForm.NewProjectComplete(p);

                Close();

            }
            else
            {
                _ = MessageBox.Show("יש למלא שם פרויקט ומספר פרויקט", "SaveAsPDF", MessageBoxButtons.OK);
            }
        }
        /// <summary>
        /// Make sure the mandatory fields are filled
        /// </summary>
        /// <returns>True if both ProjectId and ProjectName are filled</returns>
        private bool ValidateForm()
        {
            bool output = true;
            if (string.IsNullOrEmpty(txtProjectId.Text))
            {
                output = false;
            }
            if (string.IsNullOrEmpty(txtProjectName.Text))
            {
                output = true;
            }
            return output;
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void menueAdd_Click(object sender, System.EventArgs e)
        {
            TreeHelper.AddNode(tvDefaultSubFolders, mySelectedNode, "New Folder");
            
        }

        private void menuDel_Click(object sender, System.EventArgs e)
        {
            TreeHelper.DelNode(tvDefaultSubFolders, mySelectedNode);
            
        }

        private void menuRename_Click(object sender, System.EventArgs e)
        {
            TreeHelper.RenameNode(tvDefaultSubFolders, mySelectedNode);
            
        }

        private void tvDefaultSubFolders_AfterLabelEdit(object sender, NodeLabelEditEventArgs e)
        {
            if (e.Label != null)
            {
                
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
                           "אין להשתמש בתווים הבאים\n '\\', '/', ':', '*', '?', '<', '>', '|', '\"' ",
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

        private void tvDefaultSubFolders_AfterSelect(object sender, TreeViewEventArgs e)
        {
            mySelectedNode = tvDefaultSubFolders.SelectedNode;
        }
    }

}
