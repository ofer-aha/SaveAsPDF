using SaveAsPDF.Properties;
using System;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace SaveAsPDF.Helpers
{
    public static class FileFoldersHelper

    {

        /// <summary>
        /// Extention method: 
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
            if (projectNumber == null || projectNumber.Length == 0)

            {
                //Default return: root drive.
                output = Settings.Default.rootDrive;
            }
            else
            {
                if (!projectNumber.Contains("-"))
                {
                    //simple pojectid XXX or XXXX
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
            DirectoryInfo di = new DirectoryInfo(output);
            return di;
        }


        public static string SafeFileName(this string inTXT)
        {
            string pattern = @"[\/:*?""<>|]";
            //Regex rg = new Regex(pattern);

            return Regex.Replace(inTXT, pattern, "").Trim();
        }

        /// <summary>
        /// Create hidden folder
        /// </summary>
        /// <param name="folder"> string represnting the hidden folder name to create</param>
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
        /// <param name="folder">string represnting the folder name to create</param>
        public static void MkDir(string folder)
        {
            if (string.IsNullOrEmpty(folder))
            {
                throw new ArgumentNullException("MkDir:folder", "שם תקייה לא יכול להיות ריק");
            }

            // folder = folder.SafeFileName(); 

            if (Directory.Exists(folder))
            {
                int i = 0;
                if (int.TryParse(TextHelpers.GetBetween(folder, "(", ")"), out i))
                {
                    folder = folder.Replace($"({i})", $"({i+1})");
                    MkDir(folder); 
                }
                else
                {
                    folder += " (2)";
                    MkDir(folder);
                }
            }
            else
            {
                Directory.CreateDirectory(folder);
            }
        }
        /// <summary>
        /// Delete folder recursiv
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

            if (string.IsNullOrEmpty(folder))
            {
                throw new ArgumentNullException("RmDir:folder", "שם תקייה לא יכול להיות ריק");
            }

            di.MoveTo(folder);
        }
        ///// <summary>
        ///// Returns the project parent folder as string 
        ///// </summary>
        ///// <param name="pFolder"></param>
        ///// <returns></returns>
        //public static string ProjectParent( string pFolder)
        //{
        //    string[] sTemp = pFolder.Split('\\');
        //    string output = "";
        //    for (int i = 0; i < sTemp.Length - 2; i++)
        //    {
        //        output += sTemp[i]; 
        //    }
        //    return  output + "\\";
        //}
    }


}
