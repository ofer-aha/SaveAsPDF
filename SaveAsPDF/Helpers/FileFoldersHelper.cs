using Microsoft.Office.Interop.Outlook;
using SaveAsPDF.Properties;
using System;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace SaveAsPDF.Helpers
{
    public static class FileFoldersHelper

    {

        /// <summary>
        /// Extension method: 
        /// <list>
        /// Construct the full project path based on the project's number (NOT ID) 
        /// so project 1234 => j:\12\1234 
        /// </list>
        /// </summary>
        /// <param name="projectNumber"> The project name</param>
        /// <returns> string representing the project's path, default value is rootDrive</returns>
        public static DirectoryInfo ProjectFullPath(this string projectNumber)
        {
            string rootDrive = Settings.Default.rootDrive;
            string output = rootDrive;
            if (!projectNumber.SafeProjectID())

            {
                //Default return: root drive.
                output = Settings.Default.rootDrive;
            }
            else
            {
                if (!projectNumber.Contains("-"))
                {
                    //simple ProjectId XXX or XXXX
                    if (projectNumber.Length == 3)
                    {
                        //projectID: XXX = > J:\0X
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
                    //more complicated project id: XXX-X or XXX-XX or XXX-X-XX well you catch my point....
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
                    if (!Directory.Exists($"{output}\\{projectNumber}\\"))
                    {
                        // J:\XX\XXXX\XXXX-XX\
                        output += $"\\{split[0]}";
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
                        if (!Directory.Exists($"{output}\\{split[0]}-{split[1]}\\{projectNumber}\\"))
                        {

                            output += $"-{split[1]}";
                        }
                        else
                        {
                            output += $"\\{split[0]}-{split[1]}";
                        }
                    }
                }

                output += $"\\{projectNumber}\\";
            }

            return new DirectoryInfo(output); ;
        }

        //public static DirectoryInfo ProjectFullPath(this string projectNumber)
        //{
        //    if (!projectNumber.SafeProjectID())
        //    {
        //        // Default return: root drive.
        //        return new DirectoryInfo(Settings.Default.rootDrive);
        //    }

        //    string[] split = projectNumber.Split('-');
        //    string output = Settings.Default.rootDrive;


        //    if (split.Length > 0)
        //    {
        //        string prefix = split[0].Length == 3 ? $"0{split[0]}".Substring(0, 2) : split[0].Substring(0, 2);
        //        output += prefix;

        //        if (Directory.Exists($"{output}\\{projectNumber}\\"))
        //        {
        //            output += $"\\{projectNumber}\\";
        //        }

        //        if (split.Length == 3)
        //        {
        //            if (!Directory.Exists($"{output}\\{split[0]}-{split[1]}\\{projectNumber}\\"))
        //            {
        //                //output += $"-{split[1]}";
        //                output += $"\\{split[0]}-{split[1]}\\{projectNumber}\\";
        //            }
        //            else
        //            {
        //                output += $"\\{split[0]}-{split[1]}";
        //            }
        //        }
        //    }

        //    return new DirectoryInfo(output);
        //  }


        /// <summary>
        /// make sure the project pattern is right
        /// XXX XXX-X XXX-XX XXXX XXXX-X XXXX-XX and so on XXXX-XX-XX 
        /// </summary>
        /// <param name="projectID"></param>
        /// <returns></returns>
        public static Boolean SafeProjectID(this string projectID)
        {
            if (projectID == null) { return false; }

            string pattern = @"^\w{3,5}(-\w{1,3})?(-\w{1,2})?$";
            return Regex.IsMatch(projectID, pattern);
        }

        /// <summary>
        /// Created by AI 
        /// </summary>
        /// <param name="baseFolderPath"></param>
        /// <param name="desiredFolderName"></param>
        public static void CreateSafeDirectory(this DirectoryInfo directoryInfo, string desiredFolderName)
        {
            // Sanitize the folder name
            string sanitizedFolderName = new string(desiredFolderName
                .Where(c => !Path.GetInvalidPathChars().Contains(c))
                .ToArray());

            // Combine the base path with the sanitized folder name
            string fullPath = Path.Combine(directoryInfo.FullName, sanitizedFolderName);

            // Ensure the directory name is unique
            fullPath = GetUniqueDirectoryPath(fullPath);

            // Recursively create the directory
            CreateDirectoryRecursively(fullPath);
        }
        //TODO1: need to update back the treeveiw with the unique folder name 

        private static void CreateDirectoryRecursively(string path)
        {
            string parentDirectory = Path.GetDirectoryName(path);
            if (!string.IsNullOrEmpty(parentDirectory))
            {
                CreateDirectoryRecursively(parentDirectory);
            }

            if (!Directory.Exists(path))
            {
                try
                {
                    Directory.CreateDirectory(path);
                    //debugging only 
                    //MessageBox.Show("The directory was created successfully at {0}.", path);
                }
                catch (IOException ioEx)
                {
                    MessageBox.Show("An IO exception occurred: " + ioEx.Message);
                }
                catch (UnauthorizedAccessException unAuthEx)
                {
                    MessageBox.Show("UnauthorizedAccessException: " + unAuthEx.Message);
                }
                catch (ArgumentException argEx)
                {
                    MessageBox.Show("ArgumentException: " + argEx.Message);
                }
                // Additional exception handling can be added here if necessary
            }
        }

        private static string GetUniqueDirectoryPath(string path)
        {
            int counter = 1;
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
        /// <returns></returns>
        static string SafeFileName(this string folderName)
        {
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
        /// Create hidden folder. 
        /// Mainly used to create the .SaveAsPDF hidden folder 
        /// </summary>
        /// <param name="folder"> string representing the hidden folder name to create</param>
        public static void CreateHiddenFolder(this string folder)
        {
            if (!Directory.Exists(folder))
            {
                DirectoryInfo di = Directory.CreateDirectory(folder);
                di.Attributes = FileAttributes.Directory | FileAttributes.Hidden;
            }
        }

        /// <summary>
        /// Create a new folder
        /// if the folder already exists it will name it New Folder (2)... New Folder (3)  and so on. 
        /// </summary>
        /// <param name="folder">string representing the folder name to create</param>
        public static string MkDir(string folder)
        {
            string output = folder.SafeFileName();
            if (string.IsNullOrEmpty(folder))
            {
                throw new ArgumentNullException("MkDir:folder", "שם תקייה לא יכול להיות ריק");
            }

            if (Directory.Exists(folder))
            {
                int i = 0;
                if (int.TryParse(TextHelpers.GetBetween(folder, "(", ")"), out i))
                {
                    folder = folder.Replace($"({i})", $"({i + 1})");
                    output = folder;
                    MkDir(folder);
                }
                else
                {
                    output = $"{folder} (2)";
                    MkDir(output);
                }
            }
            else
            {
                Directory.CreateDirectory(folder);

            }
            return output;
        }
        /// <summary>
        /// Delete folder recursive
        /// </summary>
        /// <param name="folder">The folder to be deleted as string</param>
        public static void RmDir(string folder)
        {
            if (string.IsNullOrEmpty(folder))
            {
                throw new ArgumentNullException("RnDir:folder", "שם תקייה לא יכול להיות ריק");
            }

            string[] dirs = Directory.GetDirectories(folder);
            foreach (string dir in dirs)
            {
                RmDir(dir);
            }

            try
            {
                Directory.Delete(folder + "\\", true);
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message + "\n" + folder, "SaveAsPDF");
            }
        }
        /// <summary>
        /// Rename directory
        /// </summary>
        /// <param name="di">Directory object</param>
        /// <param name="folder">New name for the directory</param>
        /// <exception cref="ArgumentNullException"></exception>
        public static void RnDir(this DirectoryInfo di, string folder)
        {

            if (di == null)
            {
                throw new ArgumentNullException("RmDir:di", "שם תיקייה לא חוקי");
            }

            if (string.IsNullOrEmpty(folder.SafeFileName()))
            {
                throw new ArgumentNullException("RmDir:folder", "שם תקייה לא יכול להיות ריק");
            }

            di.MoveTo(folder);
        }
    }

}
