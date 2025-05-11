using System;
using System.IO;
using System.Text.RegularExpressions;

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
        public static DirectoryInfo ProjectFullPath(this string projectNumber, string rootDrive)
        {
            if (string.IsNullOrWhiteSpace(projectNumber))
                throw new ArgumentException("מספר הפרויקט לא יכול להיות ריק או שגוי.", nameof(projectNumber));

            if (string.IsNullOrWhiteSpace(rootDrive))
                throw new ArgumentException("כונן השורש לא יכול להיות ריק או שגוי.", nameof(rootDrive));

            string output = rootDrive.TrimEnd('\\') + "\\";
            projectNumber = projectNumber.Trim();

            if (!projectNumber.SafeProjectID())
                return new DirectoryInfo(output); // Return root drive if project number is invalid

            string[] split = projectNumber.Split('-');
            output += split[0].Length == 3 ? $"0{split[0]}".Substring(0, 2) : split[0].Substring(0, 2);

            if (split.Length > 1)
            {
                output += !Directory.Exists($@"{output}\{split[0]}-{split[1]}\{projectNumber}\")
                    ? $"-{split[1]}"
                    : $@"\{split[0]}-{split[1]}";
            }

            output += $@"\{projectNumber}\";
            return new DirectoryInfo(output);
        }

        /// <summary>
        /// Validates if the given project ID is in a safe format.
        /// </summary>
        public static bool SafeProjectID(this string projectID)
        {
            if (string.IsNullOrWhiteSpace(projectID))
                return false;

            string pattern = @"^[a-zA-Z0-9]{3,5}(-[a-zA-Z0-9]{1,3})?(-[a-zA-Z0-9]{1,2})?$";
            return Regex.IsMatch(projectID.Trim(), pattern);
        }

        /// <summary>
        /// Creates a directory and all its parent directories if they do not exist.
        /// </summary>
        public static void CreateDirectoryRecursively(string path)
        {
            if (string.IsNullOrWhiteSpace(path))
                throw new ArgumentException("הנתיב לא יכול להיות ריק או שגוי.", nameof(path));

            try
            {
                Directory.CreateDirectory(Path.GetFullPath(path));
            }
            catch (Exception ex) when (ex is IOException || ex is UnauthorizedAccessException || ex is ArgumentException)
            {
                XMessageBox.Show(
                    $"שגיאה ביצירת תיקיה '{path}': {ex.Message}",
                    "שגיאה ביצירת תיקיה",
                    XMessageBoxButtons.OK,
                    XMessageBoxIcon.Error,
                    XMessageAlignment.Right,
                    XMessageLanguage.Hebrew
                );
            }
        }

        /// <summary>
        /// Generates a unique directory path by appending a number to the folder name if it already exists.
        /// </summary>
        private static string GetUniqueDirectoryPath(string path)
        {
            if (string.IsNullOrWhiteSpace(path))
                throw new ArgumentException("הנתיב לא יכול להיות ריק או שגוי.", nameof(path));

            string directoryName = Path.GetFileName(path);
            string directoryPath = Path.GetDirectoryName(path) ?? throw new ArgumentException("נתיב תיקיה לא חוקי.", nameof(path));
            string uniquePath = path;
            int counter = 2;

            while (Directory.Exists(uniquePath))
                uniquePath = Path.Combine(directoryPath, $"{directoryName} ({counter++})");

            return uniquePath;
        }

        /// <summary>
        /// Sanitizes a folder name to make it safe for use in the file system.
        /// </summary>
        public static string SafeFolderName(this string folderName)
        {
            if (string.IsNullOrWhiteSpace(folderName))
                return string.Empty;

            if (Path.IsPathRooted(folderName))
            {
                string root = Path.GetPathRoot(folderName);
                string[] parts = folderName.Substring(root.Length).Split(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);

                for (int i = 0; i < parts.Length; i++)
                    parts[i] = parts[i].SafeFolderName();

                return Path.Combine(root, Path.Combine(parts));
            }

            string invalidCharsPattern = @"[\\/:*?""<>|]";
            string cleanFolderName = Regex.Replace(folderName, invalidCharsPattern, "_");

            string[] reservedNames = { "CON", "PRN", "AUX", "NUL", "COM1", "COM2", "COM3", "LPT1", "LPT2", "LPT3" };
            if (Array.Exists(reservedNames, name => string.Equals(cleanFolderName, name, StringComparison.OrdinalIgnoreCase)))
                cleanFolderName += "_";

            return cleanFolderName;
        }

        /// <summary>
        /// Creates a hidden directory at the specified path.
        /// </summary>
        public static void CreateHiddenDirectory(string folder)
        {
            try
            {
                if (Path.HasExtension(folder))
                    folder = Path.GetDirectoryName(folder);

                if (!Directory.Exists(folder))
                {
                    DirectoryInfo di = Directory.CreateDirectory(folder);
                    di.Attributes = FileAttributes.Directory | FileAttributes.Hidden;
                }
            }
            catch (Exception ex)
            {
                XMessageBox.Show(
                    $"שגיאה ביצירת תיקיה מוסתרת '{folder}': {ex.Message}",
                    "שגיאה ביצירת תיקיה מוסתרת",
                    XMessageBoxButtons.OK,
                    XMessageBoxIcon.Error,
                    XMessageAlignment.Right,
                    XMessageLanguage.Hebrew
                );
            }
        }

        /// <summary>
        /// Creates a new folder. If the folder already exists, generates a unique name.
        /// </summary>
        public static string CreateDirectory(string baseFolder)
        {
            if (string.IsNullOrEmpty(baseFolder))
            {
                XMessageBox.Show(
                    "שם התיקיה לא יכול להיות ריק.",
                    "שגיאה",
                    XMessageBoxButtons.OK,
                    XMessageBoxIcon.Error,
                    XMessageAlignment.Right,
                    XMessageLanguage.Hebrew
                );
                return baseFolder;
            }

            try
            {
                string sanitizedFolder = Path.GetFullPath(SafeFolderName(baseFolder));

                if (Directory.Exists(sanitizedFolder))
                    return Directory.CreateDirectory(GetUniqueDirectoryPath(sanitizedFolder)).FullName;

                return Directory.CreateDirectory(sanitizedFolder).FullName;
            }
            catch (Exception ex)
            {
                XMessageBox.Show(
                    $"שגיאה ביצירת תיקיה '{baseFolder}': {ex.Message}",
                    "שגיאה ביצירת תיקיה",
                    XMessageBoxButtons.OK,
                    XMessageBoxIcon.Error,
                    XMessageAlignment.Right,
                    XMessageLanguage.Hebrew
                );
                return baseFolder;
            }
        }

        /// <summary>
        /// Deletes a folder recursively.
        /// </summary>
        public static void DeleteDirectory(string folder)
        {
            if (string.IsNullOrEmpty(folder) || !Directory.Exists(folder))
            {
                XMessageBox.Show(
                    "התיקיה לא קיימת או שהשם ריק.",
                    "שגיאה",
                    XMessageBoxButtons.OK,
                    XMessageBoxIcon.Error,
                    XMessageAlignment.Right,
                    XMessageLanguage.Hebrew
                );
                return;
            }

            try
            {
                Directory.Delete(folder, true);
            }
            catch (Exception ex)
            {
                XMessageBox.Show(
                    $"שגיאה במחיקת תיקיה '{folder}': {ex.Message}",
                    "שגיאה במחיקת תיקיה",
                    XMessageBoxButtons.OK,
                    XMessageBoxIcon.Error,
                    XMessageAlignment.Right,
                    XMessageLanguage.Hebrew
                );
            }
        }

        /// <summary>
        /// Renames a directory.
        /// </summary>
        public static void RenameDirectory(this DirectoryInfo di, string folder)
        {
            if (di == null || string.IsNullOrEmpty(folder))
            {
                XMessageBox.Show(
                    "שם התיקיה לא חוקי.",
                    "שגיאה",
                    XMessageBoxButtons.OK,
                    XMessageBoxIcon.Error,
                    XMessageAlignment.Right,
                    XMessageLanguage.Hebrew
                );
                return;
            }

            try
            {
                di.MoveTo(folder);
            }
            catch (Exception ex)
            {
                XMessageBox.Show(
                    $"שגיאה בשינוי שם התיקיה '{di.FullName}': {ex.Message}",
                    "שגיאה בשינוי שם התיקיה",
                    XMessageBoxButtons.OK,
                    XMessageBoxIcon.Error,
                    XMessageAlignment.Right,
                    XMessageLanguage.Hebrew
                );
            }
        }
    }
}
