using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
//using System.Windows.Controls;
using System.Windows.Forms;

namespace SaveAsPDF.Helpers
{
    public static class ComboBoxExtensions
    {
        /// <summary>
        /// Create a combo box by loading the default list of paths from a file with ProjectID.
        /// </summary>
        /// <param name="comboBox">The combo box to load the paths into.</param>
        /// <param name="fileName">The name of the file containing the paths.</param>
        /// <param name="projectID">The ProjectID to filter or modify the paths.</param>
        public static void LoadFromFile(this ComboBox comboBox, string fileName, string projectID)
        {
            if (string.IsNullOrWhiteSpace(fileName) || string.IsNullOrWhiteSpace(projectID))
            {
                throw new ArgumentException("File name and ProjectID cannot be null or empty.");
            }

            if (!File.Exists(fileName))
            {
                throw new FileNotFoundException($"The file '{fileName}' does not exist.");
            }

            string[] list = File.ReadAllLines(fileName);
            comboBox.Items.Clear();

            foreach (string path in list)
            {
                // Replace placeholder with the ProjectID
                string modifiedPath = path.Replace(FormMain.settingsModel.ProjectRootTag, projectID);

                // Replace DateTag with the current date in "dd/MM/yyyy" format
                modifiedPath = modifiedPath.Replace(FormMain.settingsModel.DateTag, DateTime.Now.ToString("dd.MM.yyyy"));

                comboBox.Items.Add($"{FormMain.settingsModel.RootDrive}{modifiedPath}");
            }
        }

        /// <summary>
        /// Sets the ComboBox's text to a breadcrumb representation of a Windows Explorer path
        /// and populates its items with each folder segment.
        /// </summary>
        /// <param name="comboBox">The target ComboBox.</param>
        /// <param name="path">The Windows Explorer path to convert.</param>
        public static void SetBreadcrumbPath(this ComboBox comboBox, string path)
        {
            if (comboBox == null)
                throw new ArgumentNullException(nameof(comboBox));

            if (string.IsNullOrWhiteSpace(path))
                return;

            // Trim trailing backslashes to avoid empty segments.
            string trimmedPath = path.TrimEnd('\\');
            // Split the path into segments using the backslash as the delimiter.
            string[] segments = trimmedPath.Split('\\');

            // Create a breadcrumb string by joining segments with " > ".
            string breadcrumb = string.Join(" > ", segments);

            // Set the ComboBox's text to the breadcrumb representation.
            comboBox.Text = breadcrumb;

            // Clear any existing items, and then populate the ComboBox with each segment.
            comboBox.Items.Clear();
            foreach (var segment in segments)
            {
                comboBox.Items.Add(segment);
            }
        }

    }
}
