using SaveAsPDF.Models;
using SaveAsPDF.Properties;
using System.IO;
using System.Windows;

namespace SaveAsPDF.Helpers
{
    public class SettingsHelpers
    {

        static string rootDrive = @"C:\Projects\";
        static string xmlSaveAsPDFFolder = @".SaveAsPDF\";
        static string xmlProjectFile = @".SaveAsPDF_Project.xml";
        static string xmlEmployeesFile = @".SaveAsPDF_Emploeeys.xml";
        static string defaultTreeFile = @"C:\Projects\tree.fld";
        static int minAttachmentSize = 8192;
        static string dateTag = "_dateTag_";
        static int defaultFolderID = 1;
        static string projectRootTag = "_ProjectID_";
        static bool openPDF = false;
        static string lastProjects = "1000;";
        static int lastProjectsCount = 10;
        static string sProjectRootFolders = $@"{rootDrive}10\1000\";
        static string defaultSavePath = $@"{rootDrive}{sProjectRootFolders}Inbox\";



        /// <summary>
        /// Load the settings.settings to the settingsModel 
        /// </summary>
        public static SettingsModel loadSettingsToModel(SettingsModel settingsModel)
        {
            try
            {
                settingsModel.RootDrive = Settings.Default.RootDrive;
                settingsModel.XmlSaveAsPDFFolder = Settings.Default.xmlSaveAsPDFFolder;
                settingsModel.XmlProjectFile = Settings.Default.xmlProjectFile;
                settingsModel.XmlEmployeesFile = Settings.Default.xmlEmployeesFile;
                settingsModel.DefaultTreeFile = Settings.Default.DefaultTreeFile;
                settingsModel.DefaultSavePath = Settings.Default.DefaultSavePath;
                settingsModel.MinAttachmentSize = Settings.Default.MinAttachmentSize;
                settingsModel.DateTag = Settings.Default.DateTag;
                settingsModel.DefaultFolderID = Settings.Default.DefaultFolderID;
                settingsModel.ProjectRootTag = Settings.Default.ProjectRootTag;
                settingsModel.OpenPDF = Settings.Default.OpenPDF;
                settingsModel.LastProjects = Settings.Default.LastProjects;
                settingsModel.LastProjectsCount = Settings.Default.LastProjectsCount;

                settingsModel.ProjectRootFolders = string.IsNullOrEmpty(Settings.Default.sProjectRootFolders) ?
                                new DirectoryInfo(rootDrive) : new DirectoryInfo(Settings.Default.sProjectRootFolders);
            }
            catch (DirectoryNotFoundException e)
            {
                MessageBox.Show(e.Message, "SettingsHelpers:loadSettingsToModel");
            }

            return settingsModel;
        }
        /// <summary>
        /// Save the settingsModel to Settings.Settings 
        /// </summary>
        public static void saveModelToSettings(SettingsModel settingsModel)
        {
            Settings.Default.RootDrive = settingsModel.RootDrive;
            Settings.Default.xmlSaveAsPDFFolder = settingsModel.XmlSaveAsPDFFolder;
            Settings.Default.xmlProjectFile = settingsModel.XmlProjectFile;
            Settings.Default.xmlEmployeesFile = settingsModel.XmlEmployeesFile;
            Settings.Default.DefaultTreeFile = settingsModel.DefaultTreeFile;
            Settings.Default.DefaultSavePath = settingsModel.DefaultSavePath;
            Settings.Default.MinAttachmentSize = settingsModel.MinAttachmentSize;
            Settings.Default.DateTag = settingsModel.DateTag;
            Settings.Default.DefaultFolderID = settingsModel.DefaultFolderID;
            Settings.Default.ProjectRootTag = settingsModel.ProjectRootTag;
            Settings.Default.OpenPDF = settingsModel.OpenPDF;
            Settings.Default.LastProjects = settingsModel.LastProjects;
            Settings.Default.LastProjectsCount = settingsModel.LastProjectsCount;
            Settings.Default.sProjectRootFolders = settingsModel.ProjectRootFolders.ToString();


        }
        /// <summary>
        /// Ether first run or reset to defaults 
        /// </summary>
        public static void LoadDefaultSettings()
        {
            if (Settings.Default == null)
            {
                Settings.Default.RootDrive = rootDrive;
                Settings.Default.xmlSaveAsPDFFolder = xmlSaveAsPDFFolder;
                Settings.Default.xmlProjectFile = xmlProjectFile;
                Settings.Default.xmlEmployeesFile = xmlEmployeesFile;
                Settings.Default.DefaultTreeFile = defaultTreeFile;
                Settings.Default.DefaultSavePath = defaultSavePath;
                Settings.Default.MinAttachmentSize = minAttachmentSize;
                Settings.Default.DateTag = dateTag;
                Settings.Default.DefaultFolderID = defaultFolderID;
                Settings.Default.ProjectRootTag = projectRootTag;
                Settings.Default.OpenPDF = openPDF;
                Settings.Default.LastProjects = lastProjects;
                Settings.Default.LastProjectsCount = lastProjectsCount;
                Settings.Default.sProjectRootFolders = sProjectRootFolders;
            }

        }

        public static SettingsModel LoadSettingsForm(SettingsModel settingsModel)
        {
            // d
            return settingsModel;
        }

    }
}
