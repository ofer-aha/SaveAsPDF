using System;
using System.Drawing;
using System.IO;
//using System.Windows.Controls;
using System.Windows.Forms;

namespace SaveAsPDF.Helpers
{
    public static class ComboBoxExtensions
    {
        /// <summary>
        /// Loads a ComboBox with paths from a specified file, replacing placeholders with the ProjectID and current date.
        /// </summary>
        /// <param name="comboBox">The combo box to load the paths into.</param>
        /// <param name="fileName">The name of the file containing the paths.</param>
        /// <param name="projectID">The ProjectID to filter or modify the paths.</param>
        public static void LoadComboBoxWithPaths(this ComboBox comboBox, string fileName, string projectID)
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
                // Assuming FormMain.settingsModel.ProjectRootTag is a placeholder in the path

                string modifiedPath = path.Replace(FormMain.settingsModel.ProjectRootTag, FormMain.settingsModel.ProjectRootFolder.ToString());

                // Replace DateTag with the current date in "dd/MM/yyyy" format
                modifiedPath = modifiedPath.Replace(FormMain.settingsModel.DateTag, DateTime.Now.ToString("dd.MM.yyyy"));

                modifiedPath = modifiedPath.Replace("\\\\", "\\");

                // Validate the path
                if (IsValidPath(modifiedPath))
                {
                    comboBox.Items.Add(modifiedPath);
                }
                else
                {
                    // Highlight invalid paths by prefixing them with "INVALID:"
                    comboBox.Items.Add($"INVALID: {modifiedPath}");
                }
            }
        }

        /// <summary>
        /// Validates whether a given path is in a legal format.
        /// </summary>
        /// <param name="path">The path to validate.</param>
        /// <returns>True if the path is valid; otherwise, false.</returns>
        private static bool IsValidPath(string path)
        {
            try
            {
                // Check if the path is rooted and does not contain invalid characters
                return Path.IsPathRooted(path) && path.IndexOfAny(Path.GetInvalidPathChars()) == -1;
            }
            catch
            {
                // If any exception occurs, the path is considered invalid
                return false;
            }
        }
        /// <summary>
        /// Customizes the ComboBox to draw items with different colors based on their content.
        /// </summary>
        /// <param name="comboBox"></param>
        public static void CustomizeComboBox(this ComboBox comboBox)
        {
            // Set the DrawMode to OwnerDrawFixed
            comboBox.DrawMode = DrawMode.OwnerDrawFixed;

            // Handle the DrawItem event
            comboBox.DrawItem += (sender, e) =>
            {
                // Check if the index is valid
                if (e.Index < 0) return;

                // Get the item text
                string itemText = comboBox.Items[e.Index].ToString();

                // Determine the color based on the item text

                Color textColor = !Directory.Exists(itemText) ? Color.Red : Color.Black;

                // Draw the background
                e.DrawBackground();

                // Draw the text with the specified color
                using (Brush brush = new SolidBrush(textColor))
                {
                    e.Graphics.DrawString(itemText, e.Font, brush, e.Bounds);
                }

                // Draw the focus rectangle if the item is selected
                e.DrawFocusRectangle();
            };
        }

        /// <summary>
        /// Sets the ComboBox's text to a breadcrumb representation of a Windows Explorer path
        /// and populates its items with each folder segment.
        /// </summary>
        /// <param name="comboBox">The target ComboBox.</param>
        /// <param name="path">The Windows Explorer path to convert.</param>
        public static void SetBreadcrumbPath(this ComboBox comboBox)
        {
            if (comboBox == null)
                throw new ArgumentNullException(nameof(comboBox));

            // Ensure the SelectedItem is valid
            if (comboBox.SelectedItem == null || !(comboBox.SelectedItem is string path) || string.IsNullOrWhiteSpace(path))
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
