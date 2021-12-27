using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using SaveAsPDF.Helpers;
using SaveAsPDF.Models;


namespace SaveAsPDF
{
    public partial class frmNewProject : Form
    {
        private INewProjectRequester callingForm;
        private List<string> subFolfers = new List<string>();
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
    }
}
