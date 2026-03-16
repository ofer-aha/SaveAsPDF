using System;
using System.IO;
using System.Text.RegularExpressions;
using System.Collections.Generic;
using System.Text;
using System.Collections.Concurrent;

namespace SaveAsPDF.Helpers
{
    /// <summary>
    /// Provides helper methods for working with files and folders, including creating directories,
    /// sanitizing folder names, and generating unique folder paths.
    /// </summary>
    public static class FileFoldersHelper
    {
        // Cache for sanitized folder names to avoid repeated work
        private static readonly ConcurrentDictionary<string, string> _safeNameCache = new ConcurrentDictionary<string, string>(StringComparer.OrdinalIgnoreCase);

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

            projectNumber = NormalizeProjectNumber(projectNumber);
            string normalizedRoot = NormalizeRootPath(rootDrive);

            if (!projectNumber.SafeProjectID())
                return new DirectoryInfo(normalizedRoot);

            string[] split = projectNumber.Split('-');
            string firstPart = split[0];
            string level1 = GetLevel1Folder(firstPart);

            string level1Path = Path.Combine(normalizedRoot, level1);

            if (split.Length == 1)
            {
                return new DirectoryInfo(Path.Combine(level1Path, firstPart));
            }

            string flatProjectPath = Path.Combine(level1Path, projectNumber);
            string nestedBaseByFirst = Path.Combine(level1Path, firstPart);
            string nestedProjectByFirst = Path.Combine(nestedBaseByFirst, projectNumber);

            // For IDs with two hyphens, also support base folder composed of first two segments:
            // e.g. 1000-2-1 => <root>\10\1000-2\1000-2-1
            string nestedBaseByFirstTwo = split.Length > 2
                ? Path.Combine(level1Path, firstPart + "-" + split[1])
                : string.Empty;
            string nestedProjectByFirstTwo = !string.IsNullOrEmpty(nestedBaseByFirstTwo)
                ? Path.Combine(nestedBaseByFirstTwo, projectNumber)
                : string.Empty;

            // Prefer exact existing project folder first.
            if (Directory.Exists(nestedProjectByFirstTwo))
                return new DirectoryInfo(nestedProjectByFirstTwo);

            if (Directory.Exists(nestedProjectByFirst))
                return new DirectoryInfo(nestedProjectByFirst);

            if (Directory.Exists(flatProjectPath))
                return new DirectoryInfo(flatProjectPath);

            // If one of the possible bases exists, default under that base.
            if (!string.IsNullOrEmpty(nestedBaseByFirstTwo) && Directory.Exists(nestedBaseByFirstTwo))
                return new DirectoryInfo(nestedProjectByFirstTwo);

            if (Directory.Exists(nestedBaseByFirst))
                return new DirectoryInfo(nestedProjectByFirst);

            // Default for new/unknown structures: nested under first-part base folder.
            return new DirectoryInfo(nestedProjectByFirst);
        }

        private static string NormalizeProjectNumber(string projectNumber)
        {
            if (string.IsNullOrWhiteSpace(projectNumber))
                return string.Empty;

            var normalized = projectNumber.Trim();
            normalized = Regex.Replace(normalized, @"\s*-\s*", "-");
            normalized = Regex.Replace(normalized, @"\s+", string.Empty);
            return normalized;
        }

        private static string NormalizeRootPath(string rootDrive)
        {
            var root = rootDrive.Trim().Replace('/', '\\');
            if (root.Length == 2 && root[1] == ':')
                root += "\\";

            try
            {
                root = Path.GetFullPath(root);
            }
            catch
            {
            }

            // Keep drive roots rooted (J:\), never return J:
            if (root.Length == 2 && root[1] == ':')
                root += "\\";

            return root;
        }

        private static string GetLevel1Folder(string firstPart)
        {
            if (string.IsNullOrWhiteSpace(firstPart))
                return string.Empty;

            int leadingDigits = 0;
            while (leadingDigits < firstPart.Length && char.IsDigit(firstPart[leadingDigits]))
                leadingDigits++;

            if (leadingDigits >= 4)
                return firstPart.Substring(0, 2);

            // 3-digit style (and 3-digit+suffix like 123A) => 01, 02, ...
            string firstThree = firstPart.Length >= 3 ? firstPart.Substring(0, 3) : firstPart;
            return ("0" + firstThree).Substring(0, 2);
        }

        /// <summary>
        /// Validates if the given project ID is in a safe format.
        /// </summary>
        public static bool SafeProjectID(this string projectID)
        {
            if (string.IsNullOrWhiteSpace(projectID))
                return false;

            string normalized = NormalizeProjectNumber(projectID);
            string pattern = @"^\d{3,4}[a-zA-Z0-9]?(?:-[a-zA-Z0-9]{1,3})?(?:-[a-zA-Z0-9]{1,2})?$";
            return Regex.IsMatch(normalized, pattern);
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
        /// Optimized to avoid Regex and uses a cache to reduce repeated work.
        /// </summary>
        public static string SafeFolderName(this string folderName)
        {
            if (string.IsNullOrWhiteSpace(folderName))
                return string.Empty;

            // Handle rooted paths by sanitizing each segment and rejoining
            if (IsPathRootedSafe(folderName))
            {
                string root = GetPathRootSafe(folderName);
                if (!string.IsNullOrEmpty(root) && folderName.Length >= root.Length)
                {
                    var parts = folderName.Substring(root.Length)
                        .Split(new[] { Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar }, StringSplitOptions.RemoveEmptyEntries);

                    for (int i = 0; i < parts.Length; i++)
                    {
                        parts[i] = parts[i].SafeFolderName();
                    }

                    return parts.Length == 0 ? root : Path.Combine(root, Path.Combine(parts));
                }
            }

            // Use cache to avoid recomputing
            if (_safeNameCache.TryGetValue(folderName, out var cached))
                return cached;

            var invalid = Path.GetInvalidFileNameChars();
            var invalidSet = new HashSet<char>(invalid);
            var sb = new StringBuilder(folderName.Length);

            foreach (char c in folderName)
            {
                sb.Append(invalidSet.Contains(c) ? '_' : c);
            }

            string clean = sb.ToString();

            // Avoid reserved names
            string[] reservedNames = { "CON", "PRN", "AUX", "NUL", "COM1", "COM2", "COM3", "LPT1", "LPT2", "LPT3" };
            foreach (var rn in reservedNames)
            {
                if (string.Equals(clean, rn, StringComparison.OrdinalIgnoreCase))
                {
                    clean += "_";
                    break;
                }
            }

            _safeNameCache[folderName] = clean;
            return clean;
        }

        private static bool IsPathRootedSafe(string value)
        {
            try
            {
                return Path.IsPathRooted(value);
            }
            catch
            {
                return false;
            }
        }

        private static string GetPathRootSafe(string value)
        {
            try
            {
                return Path.GetPathRoot(value) ?? string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }

        /// <summary>
        /// Creates a hidden directory at the specified path.
        /// </summary>
        public static void CreateHiddenDirectory(string folder)
        {
            try
            {
                // Only treat as a file path if it has a real file extension
                // (not dot-prefixed folder names like ".SaveAsPDF")
                if (Path.HasExtension(folder) && !Path.GetFileName(folder).StartsWith("."))
                    folder = Path.GetDirectoryName(folder);

                if (string.IsNullOrWhiteSpace(folder))
                    return;

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
