using System;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace SaveAsPDF.Helpers
{
    public static class FileFoldersHelper

    {
        /// <summary>
        /// Extension method: 
        /// <list>
        /// Construct the full project path based on the projectNumber (NOT ID) 
        /// so projectNumber 1234 => j:\12\1234 
        /// </list>
        /// </summary>
        /// <param name="projectNumber"> The _projectModel name </param>
        /// <param name="rootDrive"> The default root drive </param>
        /// <returns> 
        /// string representing the _projectModel's path, default value is rootDrive
        /// </returns>
        public static DirectoryInfo ProjectFullPath(this string projectNumber, string rootDrive)
        {
            string output = rootDrive;
            projectNumber = projectNumber.Trim();

            if (!projectNumber.SafeProjectID())
            {
                //Error handling, user input was incorrect projectid format 
                //Default return: root drive.
                return new DirectoryInfo(output);
            }

            if (!projectNumber.Contains("-"))
            {
                //simple ProjectId 123 or 1234
                if (projectNumber.Length == 3)
                {
                    //projectID: 123 = > J:\01 (rootDrive+"0"+first_digit(123)) 
                    output += $"0{projectNumber}".Substring(0, 2);
                }
                else
                {
                    //projectID: XXXX => J:\XX
                    output += projectNumber.Substring(0, 2);
                }
            }
            else
            {
                //more complicated _projectModel id: XXX-X or XXX-XX or XXX-X-XX well you catch my point....
                string[] split = projectNumber.Split('-'); //split the projectid to parts

                if (split[0].Length == 3)
                {
                    output += $"0{split[0]}".Substring(0, 2);
                }
                else
                {
                    output += split[0].Substring(0, 2);
                }
                //if NOT exist   J:\XX\XXXX-XX
                if (!Directory.Exists($@"{output}\{projectNumber}\"))
                {
                    // J:\XX\XXXX\XXXX-XX\
                    output += $@"\{split[0]}";
                }
                //output = J:\
                //projectID = XXX || XXXX || XXX-X || XXXX-XX
                //split[0].substring(0,2)  = XX
                //split[0] = XXX || XXXX
                //split[1] = X || XX 
                //if exist   J:\XX\XXXX-XX   then    J:\XX\XXXX-X\       else           J:\XX\XXXX\XXXX-X

                if (split.Length == 3) //projectID looks like this XXXX-XX-XX 
                                       //   if exist   J:\XX\XXXX-XX\XXXX-XX-X\            then       
                {
                    if (!Directory.Exists($@"{output}\{split[0]}-{split[1]}\{projectNumber}\"))
                    {

                        output += $"-{split[1]}";
                    }
                    else
                    {
                        output += $@"\{split[0]}-{split[1]}";
                    }
                }


            }
            output += $@"\{projectNumber}\";
            return new DirectoryInfo(output);
        }


        /// <summary>
        /// Validates if the given projectID is in a safe format.
        /// </summary>
        /// <param name="projectID">The projectID to validate</param>
        /// <returns>True if the projectID is in a safe format, otherwise false</returns>
        public static bool SafeProjectID(this string projectID)
        {
            projectID = projectID.Trim();
            if (projectID == null) { return false; }

            string pattern = @"^\w{3,5}(-\w{1,3})?(-\w{1,2})?$";
            return Regex.IsMatch(projectID, pattern);
        }


        /// <summary>
        /// Create a directory 
        /// <list>
        /// created by AI 
        /// </list>
        /// </summary>
        /// <param name="path"></param>
        public static void CreateDirectoryRecursively(string path)
        {
            if (string.IsNullOrWhiteSpace(path))
            {
                throw new ArgumentException("Path cannot be null or empty.", nameof(path));
            }

            // Normalize the path to handle different directory separators
            path = Path.GetFullPath(path);

            // Create the directory if it doesn't exist
            if (!Directory.Exists(path))
            {
                try
                {
                    Directory.CreateDirectory(path);
                }
                catch (Exception ex) when (ex is IOException || ex is UnauthorizedAccessException)
                {
                    // Handle exceptions related to directory creation
                    Console.WriteLine($"Error creating directory '{path}': {ex.Message}");
                }
            }
        }



        //private static void CreateDirectoryRecursively(string path)
        //{
        //    string parentDirectory = Path.GetDirectoryName(path);
        //    if (!string.IsNullOrEmpty(parentDirectory))
        //    {
        //        CreateDirectoryRecursively(parentDirectory);
        //    }

        //    if (!Directory.Exists(path))
        //    {
        //        try
        //        {
        //            Directory.CreateDirectory(path);
        //            //debugging only 
        //            //MessageBox.Show("The directory was created successfully at {0}.", path);
        //        }
        //        catch (IOException ioEx)
        //        {
        //            MessageBox.Show("An IO exception occurred: " + ioEx.Message);
        //        }
        //        catch (UnauthorizedAccessException unAuthEx)
        //        {
        //            MessageBox.Show("UnauthorizedAccessException: " + unAuthEx.Message);
        //        }
        //        catch (ArgumentException argEx)
        //        {
        //            MessageBox.Show("ArgumentException: " + argEx.Message);
        //        }
        //        // Additional exception handling can be added here if necessary
        //    }
        //}

        /// <summary>
        /// Get unique folder name like "New Folder(2)" 
        /// </summary>
        /// <param name="path"></param>
        /// <returns>
        /// path
        /// </returns>
        private static string GetUniqueDirectoryPath(string path)
        {
            int counter = 2;
            string uniquePath = path;
            string directoryName = Path.GetFileName(path);
            string directoryPath = Path.GetDirectoryName(path);

            while (Directory.Exists(uniquePath))
            {
                // If the directory exists, append a number to make it unique
                uniquePath = Path.Combine(directoryPath, $"{directoryName} ({counter++})");
            }

            return uniquePath;
        }
        /// <summary>
        ///  Define a regex pattern to match invalid characters(e.g., backslash, colon, asterisk, question mark, double quotes, angle brackets, and pipe).
        ///  Replace these invalid characters with underscores.
        ///  Check if the sanitized folder name matches any system reserved names (e.g., “CON,” “PRN,” etc.). If it does, append an underscore to avoid conflicts.
        ///  Ensure the resulting folder name is safe for use.
        /// </summary>
        /// <param name="folderName"></param>
        /// <returns>
        /// A safe file-system folder name
        /// </returns>
        public static string SafeFolderName(this string folderName)
        {
            if (string.IsNullOrEmpty(folderName))
            {
                return folderName;
            }
            // Define the regex pattern to match invalid characters
            string invalidCharsPattern = @"[\\/:*?""<>|]";

            // Replace invalid characters with underscores
            string cleanFolderName = Regex.Replace(folderName, invalidCharsPattern, "_");
            // Check if the sanitized folder name matches any system reserved names
            string[] reservedNames = { "CON", "PRN", "AUX", "NUL", "COM1", "COM2", "COM3", "LPT1", "LPT2", "LPT3" };
            foreach (string reservedName in reservedNames)
            {
                if (string.Equals(cleanFolderName, reservedName, StringComparison.OrdinalIgnoreCase))
                {
                    // Append an underscore to avoid conflicts with reserved names
                    cleanFolderName += "_";
                    break;
                }
            }
            return cleanFolderName;
        }


        /// <summary>
        /// <list>
        /// Extension method to create hidden folders.
        /// </list>
        /// Mainly used to create the <b><u>.SaveAsPDF </u></b> hidden folder 
        /// use internally - no need to verify input
        /// </summary>
        /// <param name="folder"> string representing the hidden folder name to create </param>
        /// <param name="defName"> default folder name. When left empty the default is to create the .SaveAsPDF folder </param>
        /// 
        public static void CreateHiddenDirectory(this string folder, string defName = null)
        {
            if (Path.HasExtension(folder))
            {
                folder = Path.GetDirectoryName(folder);
            }

            if (!string.IsNullOrEmpty(defName))
            {
                if (!Directory.Exists($@"{folder}\{defName}"))
                {
                    DirectoryInfo di = Directory.CreateDirectory($@"{folder}\{defName}");
                    di.Attributes = FileAttributes.Directory | FileAttributes.Hidden;
                }
            }
            else
            {
                try
                {
                    DirectoryInfo di = Directory.CreateDirectory(folder);
                    di.Attributes = FileAttributes.Directory | FileAttributes.Hidden;
                }
                catch (Exception e)
                {
                    MessageBox.Show(e.Message, "FileFolderHelper:CreatHiddenFolde");
                }
            }
        }

        /// <summary>
        /// Create a new folder
        /// if the folder already exists it will name it New Folder (2)... New Folder (3)  and so on. 
        /// </summary>
        /// <param name="folder">string representing the folder name to create</param>
        public static string CreateDirectory(string baseFolder)
        {
            if (string.IsNullOrEmpty(baseFolder))
            {
                MessageBox.Show("שם תיקיה לא יכול להיות ריק");
                return baseFolder;
            }
            // Remove invalid characters from the folder name
            //string sanitizedFolder = SanitizeFolderName(baseFolder);
            string sanitizedFolder = baseFolder;
            // Check if the directory already exists
            if (Directory.Exists(sanitizedFolder))
            {
                // Create the unique directory
                string uniqueFolder = GetUniqueDirectoryPath(sanitizedFolder);

                Directory.CreateDirectory(uniqueFolder);
                return uniqueFolder;
            }
            else
            {
                // Create the directory
                Directory.CreateDirectory(sanitizedFolder);
                return sanitizedFolder;
            }
        }

        /// <summary>
        /// Delete folder recursive
        /// </summary>
        /// <param name="folder">The folder to be deleted as string</param>
        public static void DeleteDirectory(string folder)
        {
            if (string.IsNullOrEmpty(folder) || !Directory.Exists(folder))
            {
                MessageBox.Show("תיקיה לא קיימת או שם ריק");
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
                MessageBox.Show(e.Message + "\n" + folder, "SaveAsPDF");
            }
        }
        /// <summary>
        /// Extension method to Rename a directory
        /// </summary>
        /// <param name="di">Directory object</param>
        /// <param name="folder">New name for the directory</param>
        public static void RenameDirectory(this DirectoryInfo di, string folder)
        {
            if (di == null || string.IsNullOrEmpty(folder))
            {
                MessageBox.Show("שם תיקייה לא חוקי");
                return;
            }
            else
            {
                di.MoveTo(folder);
            }
        }

    }
}
