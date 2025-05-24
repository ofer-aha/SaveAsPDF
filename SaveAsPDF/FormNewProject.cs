using SaveAsPDF.Helpers;
using SaveAsPDF.Models;
using System;
using System.Collections.Generic;
using System.Windows.Forms;

namespace SaveAsPDF
{
    /// <summary>
    /// Represents the form for creating a new project, including project details and default sub-folders.
    /// </summary>
    public partial class FormNewProject : Form
    {
        /// <summary>
        /// The form that requested the creation of a new project.
        /// </summary>
        private readonly INewProjectRequester callingForm;

        /// <summary>
        /// The list of sub-folder paths for the new project.
        /// </summary>
        private List<string> _subFolfers = new List<string>();

        /// <summary>
        /// The currently selected node in the sub-folders tree view.
        /// </summary>
        private TreeNode _mySelectedNode;

        /// <summary>
        /// Initializes a new instance of the <see cref="FormNewProject"/> class.
        /// </summary>
        /// <param name="caller">The form that requested the new project.</param>
        public FormNewProject(INewProjectRequester caller)
        {
            InitializeComponent();
            callingForm = caller;
            txtProjectNotes.EnableContextMenu();
            txtProjectId.EnableContextMenu();
            txtProjectName.EnableContextMenu();
            // Optionally initialize default sub-folders here if needed
            LoadSettings();
        }

        /// <summary>
        /// Loads the settings for the form, such as default sub-folders.
        /// </summary>
        private void LoadSettings()
        {
            // Removed tvDefaultSubFolders.LoadFromList(); as it does not exist.
            // If you need to initialize nodes, do it here, e.g.:
            // tvDefaultSubFolders.Nodes.Clear();
            // tvDefaultSubFolders.Nodes.Add("תיקיה ראשית");
        }

        /// <summary>
        /// Handles the click event for the "Create New Project" button.
        /// Validates the form, creates the project model, creates sub-folders, and notifies the calling form.
        /// </summary>
        private void btmNewProject_Click(object sender, EventArgs e)
        {
            if (!ValidateForm())
            {
                XMessageBox.Show(
                    "יש למלא שם פרויקט ומספר פרויקט",
                    "SaveAsPDF",
                    XMessageBoxButtons.OK,
                    XMessageBoxIcon.Warning,
                    XMessageAlignment.Right,
                    XMessageLanguage.Hebrew
                );
                return;
            }

            var projectModel = new ProjectModel
            {
                ProjectName = txtProjectName.Text,
                ProjectNumber = txtProjectId.Text.Trim(),
                ProjectNotes = txtProjectNotes.Text
            };

            if (tvDefaultSubFolders.Nodes.Count > 0)
            {
                _subFolfers = TreeHelpers.ListNodesPath(tvDefaultSubFolders.Nodes[0]);
                foreach (var subFolder in _subFolfers)
                {
                    FileFoldersHelper.CreateDirectory(subFolder);
                }
            }

            callingForm.NewProjectComplete(projectModel);
            Close();
        }

        /// <summary>
        /// Validates that the mandatory fields (ProjectId and ProjectName) are filled.
        /// </summary>
        /// <returns>True if both ProjectId and ProjectName are filled; otherwise, false.</returns>
        private bool ValidateForm()
        {
            return !string.IsNullOrWhiteSpace(txtProjectId.Text) && !string.IsNullOrWhiteSpace(txtProjectName.Text);
        }

        /// <summary>
        /// Handles the click event for the "Cancel" button and closes the form.
        /// </summary>
        private void btnCancel_Click(object sender, EventArgs e)
        {
            Close();
        }

        /// <summary>
        /// Handles the click event for the "Add" menu item to add a new sub-folder node.
        /// </summary>
        private void menueAdd_Click(object sender, EventArgs e)
        {
            TreeHelpers.AddNode(tvDefaultSubFolders, _mySelectedNode);
        }

        /// <summary>
        /// Handles the click event for the "Delete" menu item to delete the selected sub-folder node.
        /// </summary>
        private void menuDel_Click(object sender, EventArgs e)
        {
            TreeHelpers.DeleteNode(tvDefaultSubFolders);
        }

        /// <summary>
        /// Handles the click event for the "Rename" menu item to rename the selected sub-folder node.
        /// </summary>
        private void menuRename_Click(object sender, EventArgs e)
        {
            TreeHelpers.RenameNode(tvDefaultSubFolders, _mySelectedNode);
        }

        /// <summary>
        /// Handles the AfterLabelEdit event for the sub-folders tree view.
        /// Validates the new label and displays an error if invalid.
        /// </summary>
        private void tvDefaultSubFolders_AfterLabelEdit(object sender, NodeLabelEditEventArgs e)
        {
            if (e.Label == null)
            {
                e.CancelEdit = true;
                XMessageBox.Show(
                    "שם לא חוקי.\nלא ניתן ליצור שם ריק. חובה תו אחד לפחות",
                    "עריכת שם",
                    XMessageBoxButtons.OK,
                    XMessageBoxIcon.Warning,
                    XMessageAlignment.Right,
                    XMessageLanguage.Hebrew
                );
                e.Node?.BeginEdit();
                return;
            }

            if (e.Label.Length == 0 || e.Label.IndexOfAny(new[] { '\\', '/', ':', '*', '?', '<', '>', '|', '"' }) != -1)
            {
                e.CancelEdit = true;
                XMessageBox.Show(
                    e.Label.Length == 0
                        ? "שם לא חוקי.\nלא ניתן ליצור שם ריק. חובה תו אחד לפחות"
                        : "שם לא חוקי.\nאין להשתמש בתווים הבאים\n '\\', '/', ':', '*', '?', '<', '>', '|', '\"' ",
                    "עריכת שם",
                    XMessageBoxButtons.OK,
                    XMessageBoxIcon.Warning,
                    XMessageAlignment.Right,
                    XMessageLanguage.Hebrew
                );
                e.Node?.BeginEdit();
            }
        }

        /// <summary>
        /// Handles the AfterSelect event for the sub-folders tree view.
        /// Updates the currently selected node.
        /// </summary>
        private void tvDefaultSubFolders_AfterSelect(object sender, TreeViewEventArgs e)
        {
            _mySelectedNode = tvDefaultSubFolders.SelectedNode;
        }

        /// <summary>
        /// Handles the Load event for the form.
        /// </summary>
        private void FormNewProject_Load(object sender, EventArgs e)
        {
            // Optionally set default values here
        }

        /// <summary>
        /// Handles the Validating event for the ProjectId textbox.
        /// Validates the project ID and updates the UI accordingly.
        /// </summary>
        private void txtProjectId_Validating(object sender, System.ComponentModel.CancelEventArgs e)
        {
            if (!FileFoldersHelper.SafeProjectID(txtProjectId.Text))
            {
                e.Cancel = true;
                txtProjectId.Select(0, txtProjectId.Text.Length);
                txtProjectId.BackColor = System.Drawing.Color.Red;
                toolStripStatusLabel.Text = "מספר פרויקט לא חוקי";
            }
        }

        /// <summary>
        /// Handles the Validated event for the ProjectId textbox.
        /// Resets the UI to indicate valid input.
        /// </summary>
        private void txtProjectId_Validated(object sender, EventArgs e)
        {
            txtProjectId.BackColor = System.Drawing.Color.White;
            toolStripStatusLabel.Text = string.Empty;
        }

        /// <summary>
        /// Handles the Validating event for the ProjectName textbox.
        /// Validates the project name and updates the UI accordingly.
        /// </summary>
        private void txtProjectName_Validating(object sender, System.ComponentModel.CancelEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtProjectName.Text))
            {
                e.Cancel = true;
                txtProjectName.Select(0, txtProjectName.Text.Length);
                txtProjectName.BackColor = System.Drawing.Color.Red;
                toolStripStatusLabel.Text = "שם פרויקט לא חוקי";
            }
        }

        /// <summary>
        /// Handles the Validated event for the ProjectName textbox.
        /// Resets the UI to indicate valid input.
        /// </summary>
        private void txtProjectName_Validated(object sender, EventArgs e)
        {
            txtProjectName.BackColor = System.Drawing.Color.White;
            toolStripStatusLabel.Text = string.Empty;
        }
    }
}