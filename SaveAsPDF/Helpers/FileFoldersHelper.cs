using System;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace SaveAsPDF.Helpers
{
    /// <summary>
    /// Provides helper methods for working with files and folders, including creating directories,
    /// sanitizing folder names, and generating unique folder paths.
    /// </summary>
    public static class FileFoldersHelper
    {
        /// <summary>
        /// Constructs the full project path based on the project number.
        /// For example, project number 1234 => j:\12\1234.
        /// </summary>
        /// <param name="projectNumber">The project number.</param>
        /// <param name="rootDrive">The root drive where the project resides.</param>
        /// <returns>A <see cref="DirectoryInfo"/> object representing the full project path.</returns>
        public static DirectoryInfo ProjectFullPath(this string projectNumber, string rootDrive)
        {
            if (string.IsNullOrWhiteSpace(projectNumber))
            {
                throw new ArgumentException("מספר הפרויקט לא יכול להיות ריק או שגוי.", nameof(projectNumber));
            }

            if (string.IsNullOrWhiteSpace(rootDrive))
            {
                throw new ArgumentException("כונן השורש לא יכול להיות ריק או שגוי.", nameof(rootDrive));
            }

            string output = rootDrive.TrimEnd('\\') + "\\"; // Ensure rootDrive ends with a backslash
            projectNumber = projectNumber.Trim();

            if (!projectNumber.SafeProjectID())
            {
                // Return the root drive if the project number is invalid
                return new DirectoryInfo(output);
            }

            if (!projectNumber.Contains("-"))
            {
                // Simple project number (e.g., 123 or 1234)
                output += projectNumber.Length == 3
                    ? $"0{projectNumber}".Substring(0, 2)
                    : projectNumber.Substring(0, 2);
            }
            else
            {
                // Complex project number (e.g., XXX-X or XXXX-XX)
                string[] split = projectNumber.Split('-');

                output += split[0].Length == 3
                    ? $"0{split[0]}".Substring(0, 2)
                    : split[0].Substring(0, 2);

                if (!Directory.Exists($@"{output}\{projectNumber}\"))
                {
                    output += $@"\{split[0]}";
                }

                if (split.Length == 3)
                {
                    output += !Directory.Exists($@"{output}\{split[0]}-{split[1]}\{projectNumber}\")
                        ? $"-{split[1]}"
                        : $@"\{split[0]}-{split[1]}";
                }
            }

            output += $@"\{projectNumber}\";
            return new DirectoryInfo(output);
        }

        /// <summary>
        /// Validates if the given project ID is in a safe format.
        /// A valid project ID must match the following pattern:
        /// - 3 to 5 alphanumeric characters, optionally followed by:
        ///   - A hyphen and 1 to 3 alphanumeric characters, optionally followed by:
        ///   - Another hyphen and 1 to 2 alphanumeric characters.
        /// </summary>
        /// <param name="projectID">The project ID to validate.</param>
        /// <returns>True if the project ID is in a safe format, otherwise false.</returns>
        public static bool SafeProjectID(this string projectID)
        {
            if (string.IsNullOrWhiteSpace(projectID))
            {
                return false;
            }

            projectID = projectID.Trim();
            string pattern = @"^[a-zA-Z0-9]{3,5}(-[a-zA-Z0-9]{1,3})?(-[a-zA-Z0-9]{1,2})?$";
            return Regex.IsMatch(projectID, pattern);
        }

        /// <summary>
        /// Creates a directory and all its parent directories if they do not exist.
        /// </summary>
        /// <param name="path">The full path of the directory to create.</param>
        public static void CreateDirectoryRecursively(string path)
        {
            if (string.IsNullOrWhiteSpace(path))
            {
                throw new ArgumentException("הנתיב לא יכול להיות ריק או שגוי.", nameof(path));
            }

            try
            {
                path = Path.GetFullPath(path);
                Directory.CreateDirectory(path);
            }
            catch (Exception ex) when (ex is IOException || ex is UnauthorizedAccessException || ex is ArgumentException)
            {
                MessageBox.Show($"שגיאה ביצירת תיקיה '{path}': {ex.Message}", "SaveAsPDF:CreateDirectoryRecursively", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Generates a unique directory path by appending a number to the folder name if it already exists.
        /// For example, if "New Folder" exists, it will create "New Folder (2)", "New Folder (3)", and so on.
        /// </summary>
        /// <param name="path">The base path of the directory to check for uniqueness.</param>
        /// <returns>A unique directory path.</returns>
        private static string GetUniqueDirectoryPath(string path)
        {
            if (string.IsNullOrWhiteSpace(path))
            {
                throw new ArgumentException("הנתיב לא יכול להיות ריק או שגוי.", nameof(path));
            }

            string directoryName = Path.GetFileName(path);
            string directoryPath = Path.GetDirectoryName(path);
            string uniquePath = path;
            int counter = 2;

            if (string.IsNullOrEmpty(directoryPath))
            {
                throw new ArgumentException("נתיב תיקיה לא חוקי.", nameof(path));
            }

            while (Directory.Exists(uniquePath))
            {
                uniquePath = Path.Combine(directoryPath, $"{directoryName} ({counter++})");
            }

            return uniquePath;
        }

        /// <summary>
        /// Sanitizes a folder name to make it safe for use in the file system.
        /// - Replaces invalid characters (e.g., \ / : * ? " < > |) with underscores.
        /// - Appends an underscore if the name matches a reserved system name (e.g., CON, PRN).
        /// </summary>
        /// <param name="folderName">The folder name to sanitize.</param>
        /// <returns>A sanitized folder name that is safe for use in the file system.</returns>
        public static string SafeFolderName(this string folderName)
        {
            if (string.IsNullOrWhiteSpace(folderName))
            {
                throw new ArgumentException("שם התיקיה לא יכול להיות ריק או שגוי.", nameof(folderName));
            }

            string invalidCharsPattern = @"[\\/:*?""<>|]";
            string cleanFolderName = Regex.Replace(folderName, invalidCharsPattern, "_");

            string[] reservedNames = { "CON", "PRN", "AUX", "NUL", "COM1", "COM2", "COM3", "LPT1", "LPT2", "LPT3" };
            foreach (string reservedName in reservedNames)
            {
                if (string.Equals(cleanFolderName, reservedName, StringComparison.OrdinalIgnoreCase))
                {
                    cleanFolderName += "_";
                    break;
                }
            }

            return cleanFolderName;
        }

        /// <summary>
        /// Creates a hidden directory at the specified path.
        /// If a default folder name is provided, it appends it to the path.
        /// </summary>
        /// <param name="folder">The base path where the hidden folder will be created.</param>
        /// <param name="defName">Optional default folder name to append to the path.</param>
        public static void CreateHiddenDirectory(this string folder, string defName = null)
        {
            try
            {
                if (Path.HasExtension(folder))
                {
                    folder = Path.GetDirectoryName(folder);
                }

                if (!string.IsNullOrEmpty(defName))
                {
                    folder = Path.Combine(folder, defName);
                }

                if (!Directory.Exists(folder))
                {
                    DirectoryInfo di = Directory.CreateDirectory(folder);
                    di.Attributes = FileAttributes.Directory | FileAttributes.Hidden;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"שגיאה ביצירת תיקיה מוסתרת '{folder}': {ex.Message}", "FileFoldersHelper:CreateHiddenDirectory", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Creates a new folder.
        /// If the folder already exists, it generates a unique name by appending a number (e.g., "New Folder (2)", "New Folder (3)").
        /// </summary>
        /// <param name="baseFolder">The base path of the folder to create.</param>
        /// <returns>The path of the created folder.</returns>
        public static string CreateDirectory(string baseFolder)
        {
            if (string.IsNullOrEmpty(baseFolder))
            {
                MessageBox.Show("שם התיקיה לא יכול להיות ריק.", "שגיאה", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return baseFolder;
            }

            string sanitizedFolder = SafeFolderName(baseFolder);

            if (Directory.Exists(sanitizedFolder))
            {
                string uniqueFolder = GetUniqueDirectoryPath(sanitizedFolder);
                Directory.CreateDirectory(uniqueFolder);
                return uniqueFolder;
            }
            else
            {
                Directory.CreateDirectory(sanitizedFolder);
                return sanitizedFolder;
            }
        }

        /// <summary>
        /// Deletes a folder recursively.
        /// </summary>
        /// <param name="folder">The folder to delete.</param>
        public static void DeleteDirectory(string folder)
        {
            if (string.IsNullOrEmpty(folder) || !Directory.Exists(folder))
            {
                MessageBox.Show("התיקיה לא קיימת או שהשם ריק.", "שגיאה", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            string[] dirs = Directory.GetDirectories(folder);
            foreach (string dir in dirs)
            {
                DeleteDirectory(dir);
            }
            try
            {
                Directory.Delete($@"{folder}\", true);
            }
            catch (Exception e)
            {
                MessageBox.Show($"שגיאה במחיקת תיקיה '{folder}': {e.Message}", "SaveAsPDF", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Renames a directory.
        /// </summary>
        /// <param name="di">The <see cref="DirectoryInfo"/> object representing the directory to rename.</param>
        /// <param name="folder">The new name for the directory.</param>
        public static void RenameDirectory(this DirectoryInfo di, string folder)
        {
            if (di == null || string.IsNullOrEmpty(folder))
            {
                MessageBox.Show("שם התיקיה לא חוקי.", "שגיאה", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            else
            {
                di.MoveTo(folder);
            }
        }
    }
}
