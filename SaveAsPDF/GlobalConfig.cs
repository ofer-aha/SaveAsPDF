using System;
using System.Configuration;
using System.Windows.Forms;

namespace SaveAsPDF
{
    public static class GlobalConfig
    {
        /// <summary>
        /// get the projects root directory from app.config "J:\" 
        /// </summary>
        public static string rootDrive = $"{ ConfigurationManager.AppSettings["rootDrive"]}";
        /// <summary>
        /// the hidden subfolder containing all SaveAsPDF XMP files
        /// </summary>
        public static string xmlSaveAsPdfFolder = $"{ ConfigurationManager.AppSettings["xmlSaveAsPdfFolder"]}";
        /// <summary>
        /// Get the project XML file name
        /// </summary>
        public  static string xmlProjectFile = $"{ConfigurationManager.AppSettings["xmlProjectFile"]}";
        /// <summary>
        /// Get the employees XML file name
        /// </summary>
        public static string xmlEmploeeysFile = $"{ConfigurationManager.AppSettings["xmlEmploeeysFile"]}";


    }



}