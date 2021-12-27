using System;
using System.Configuration;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Text.RegularExpressions;
using SaveAsPDF.Models;
using System.Windows.Forms;
using System.Runtime.Remoting.Messaging;
using SaveAsPDF.Properties;

namespace SaveAsPDF.Helpers
{
    public static class FileFoldersHelper

    {
        
        /// <summary>
        /// Extention method: 
        /// Construct the full project path based on the project's number (NOT ID) 
        /// so project 1234 => j:\12\1234 
        /// </summary>
        /// <param name="projectNumber"> The project name</param>
        /// <returns> string representing the project's path, default value is rootDrive</returns>
        public static string ProjectFullPath(this string projectNumber)
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

                            output +=  $"-{split[1]}";
                        }
                        else
                        {
                            output += $"\\{split[0]}-{split[1]}";
                        }
                    }
                }

                output += $"\\{projectNumber}\\";
            }
            return output;
        }


        public static string SafeFileName( this string inTXT)
        {
            string pattern = @"[\/:*?""<>|]";
            //Regex rg = new Regex(pattern);
            
            return Regex.Replace(inTXT, pattern, "").Trim();
        }

        /// <summary>
        /// Create hidden folder
        /// </summary>
        /// <param name="folder"> string represnting the hidden folder name to create</param>
        public static void CreateHiddenFolder( string folder)
        {
            if (!Directory.Exists(folder))
            {
                DirectoryInfo di = Directory.CreateDirectory(folder);
                di.Attributes = FileAttributes.Directory | FileAttributes.Hidden;
            }
        }

        /// <summary>
        /// Create a new folder
        /// if the folder already exists it will name it Folder (2)... Folder (2) (2) and so on. 
        /// </summary>
        /// <param name="folder">string represnting the folder name to create</param>
        public static void MkDir( string folder)
        {
            
            if (Directory.Exists(folder))
            {
                int i=1;
                folder += $" ({i + 1})";
                MkDir(folder); 
            }
            else
            {
                Directory.CreateDirectory(folder);
            }
        }
        /// <summary>
        /// Delete folder 
        /// </summary>
        /// <param name="folder"> the folder to be deleted as string</param>
        public static void RmDir (string folder)
        {
            try
            {
                Directory.Delete (folder);
            }
            catch (Exception e)
            {

                MessageBox.Show(e.Message, "SaveAsPDF");
            }
        }

        public static void RnDir(this DirectoryInfo di, string folder)
        {
            if (di == null)
            {
                throw new ArgumentNullException("di", "שם תיקייה לא חוקי");
            }
            if (string.IsNullOrEmpty(folder))
            {
                throw new ArgumentNullException("שם תקייה לא יכול להיות ריק", "folder");
            }
            di.MoveTo(Path.Combine(di.Parent.FullName, folder));
        }
        
    }


}
