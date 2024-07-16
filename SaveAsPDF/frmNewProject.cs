// Ignore Spelling: frm מפשחה הכל יש לבחור הודעות דואר אלקטרוני בלבד אימייל הסר הכול שם קובץ גודל יש לבחור הודעות דואר אלקטרוני בלבד ההודעה נשמרה ב  תאריך  שמירה  שם הפרויקט  מס פרויקט  הערות  שם משתמש בחר  הסר  מספר פרויקט כפי שמופיע במסטרפלן שם לא חוקי  אין להשתמש בתווים הבאים  עריכת שם שם לא חוקי לא ניתן ליצור שם ריק חובה תו אחד לפחות עריכת שם מספר פרויקט לא חוקי
using SaveAsPDF.Helpers;
using SaveAsPDF.Models;
using System;
using System.Collections.Generic;
using System.Windows.Forms;


namespace SaveAsPDF
{
    public partial class frmNewProject : Form
    {
        private readonly INewProjectRequester callingForm;
        private List<string> _subFolfers = new List<string>();
        private TreeNode _mySelectedNode;
        public frmNewProject(INewProjectRequester caller)
        {
            InitializeComponent();
            txtProjectNotes.EnableContextMenu();
            txtProjectId.EnableContextMenu();
            txtProjectName.EnableContextMenu();
            LoadSettings();
        }

        /// <summary>
        /// Loads the settings for the form.
        /// </summary>
        private void LoadSettings()
        {
            tvDefaultSubFolders.LoadTreeViewFromList();
        }

        private void btmNewProject_Click(object sender, EventArgs e)
        {
            //populate ProjectModel
            if (ValidateForm())
            {
                ProjectModel projectModel = new ProjectModel
                {
                    ProjectName = txtProjectName.Text,
                    ProjectNumber = txtProjectId.Text.Trim(),
                    ProjectNotes = txtProjectNotes.Text
                };

                //create the _projectModel's sub-folders 

                _subFolfers = TreeHelpers.ListNodesPath(tvDefaultSubFolders.Nodes[0]);
                foreach (string subFolder in _subFolfers)
                {
                    FileFoldersHelper.MkDir(subFolder);
                }

                //call the calling form to add the new project
                //TODO1: initialize the projectModel 
                callingForm.NewProjectComplete(projectModel);

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
            tvDefaultSubFolders.AddNode(_mySelectedNode);

        }

        private void menuDel_Click(object sender, System.EventArgs e)
        {
            tvDefaultSubFolders.DelNode(_mySelectedNode);

        }

        private void menuRename_Click(object sender, System.EventArgs e)
        {
            tvDefaultSubFolders.RenameNode(_mySelectedNode);
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
            _mySelectedNode = tvDefaultSubFolders.SelectedNode;
        }

        private void frmNewProject_Load(object sender, EventArgs e)
        {
            //TODO: default values? 
        }

        private void txtProjectId_Validating(object sender, System.ComponentModel.CancelEventArgs e)
        {
            if (!FileFoldersHelper.SafeProjectID(txtProjectId.Text))
            {
                e.Cancel = true;
                txtProjectId.Select(0, txtProjectId.Text.Length);
                txtProjectId.BackColor = System.Drawing.Color.Red;
                this.toolStripStatusLabel.Text = "מספר פרויקט לא חוקי";
            }
        }

        private void txtProjectId_Validated(object sender, EventArgs e)
        {
            txtProjectId.BackColor = System.Drawing.Color.White;
            toolStripStatusLabel.Text = string.Empty;
        }

        private void txtProjectName_Validating(object sender, System.ComponentModel.CancelEventArgs e)
        {
            if (txtProjectName.Text.Trim().Length == 0)
            {
                e.Cancel = true;
                txtProjectName.Select(0, txtProjectId.Text.Length);
                txtProjectName.BackColor = System.Drawing.Color.Red;
                toolStripStatusLabel.Text = "שם פרויקט לא חוקי";
            }
        }

        private void txtProjectName_Validated(object sender, EventArgs e)
        {
            txtProjectName.BackColor = System.Drawing.Color.White;
            toolStripStatusLabel.Text = string.Empty;
        }
    }

}
