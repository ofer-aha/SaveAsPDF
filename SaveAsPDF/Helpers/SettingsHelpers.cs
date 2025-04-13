using SaveAsPDF.Models;
using SaveAsPDF.Properties;
using System;
using System.IO;
using System.Windows;

namespace SaveAsPDF.Helpers
{
    [Serializable]
    public static class SettingsHelpers
    {
        //readonly for default settings
        static public readonly string rootDrive = @"C:\Projects\";
        static public readonly string xmlSaveAsPDFFolder = @".SaveAsPDF\";
        static public readonly string xmlProjectFile = @".SaveAsPDF_Project.xml";
        static public readonly string xmlEmployeesFile = @".SaveAsPDF_Emploeeys.xml";
        static public readonly string defaultTreeFile = @"C:\Projects\tree.fld";
        static public readonly int minAttachmentSize = 8192;
        static public readonly string dateTag = "_תאריך_";
        static public readonly int defaultFolderID = 1;
        static public readonly string projectRootTag = "_מספר_פרויקט_";
        static public readonly bool openPDF = false;
        static public readonly string lastProjects = "1000;";
        static public readonly int lastProjectsCount = 10;
        static public readonly string sProjectRootFolders = $@"{rootDrive}10\1000\";
        static public readonly string defaultSavePath = $@"{rootDrive}{sProjectRootFolders}Inbox\";

        /// <summary>
        /// Load the settings.settings to the _settingsModel 
        /// </summary>
        /// <param name="settingsModel"><see cref="SettingsModel"/> object</param>
        /// <returns><see cref="SettingsModel"/> object</returns>
        public static SettingsModel LoadSettingsToModel(SettingsModel settingsModel)
        {
            try
            {
                // Load the settings from Settings.Default
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

                // Deserialize the SettingsModelJson
                //SettingsModel settingsModelJson = Settings.Default.SettingsModelJson;
                //if (settingsModelJson != null)
                //{
                //    settingsModel = JsonConvert.DeserializeObject<SettingsModel>(settingsModelJson);
                //    //settingsModel = JsonSerializer(settingsModelJson);
                //}

                // Set the ProjectRootFolder
                settingsModel.ProjectRootFolder = string.IsNullOrEmpty(Settings.Default.sProjectRootFolders) ?
                    new DirectoryInfo(rootDrive) : new DirectoryInfo(Settings.Default.sProjectRootFolders);
            }
            catch (System.Exception e)
            {
                MessageBox.Show(e.Message, "SettingsHelpers:LoadSettingsToModel");
            }

            return settingsModel;
        }

        /// <summary>
        /// Save the _settingsModel to Settings.Settings 
        /// </summary>
        /// <param name="settingsModel"><see cref="SettingsModel"/> object</param>
        /// 
        public static void SaveModelToSettings(SettingsModel settingsModel)
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
            Settings.Default.sProjectRootFolders = settingsModel.ProjectRootFolder.ToString();
            Settings.Default.Save();
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
                Settings.Default.MaxProjectCount = lastProjectsCount;
                Settings.Default.sProjectRootFolders = sProjectRootFolders;
            }
        }


        /// <summary>
        /// Load the settingsModel from Settings.Settings. parameterless overload.  
        /// </summary>
        /// <returns><see cref="SettingsModel"/> object </returns>
        public static SettingsModel LoadProjectSettings()
        {
            return LoadProjectSettings(string.Empty);
        }
        /// <summary>
        /// Load the settingsModel from Settings.Settings 
        /// </summary>
        /// <param name="projectID"><see cref="string"/> projrctID</param>
        /// <returns><see cref="SettingsModel"/> object </returns>
        public static SettingsModel LoadProjectSettings(this string projectID)
        {
            SettingsModel settingsModel = new SettingsModel();

            settingsModel.RootDrive = Settings.Default.RootDrive; // j:\

            // Folder tags
            settingsModel.DateTag = Settings.Default.DateTag;
            settingsModel.ProjectRootTag = Settings.Default.ProjectRootTag;

            // More settings
            settingsModel.DefaultFolderID = Settings.Default.DefaultFolderID;
            settingsModel.MinAttachmentSize = Settings.Default.MinAttachmentSize;
            settingsModel.DefaultTreeFile = Settings.Default.DefaultTreeFile;
            settingsModel.OpenPDF = Settings.Default.OpenPDF;


            if (!string.IsNullOrEmpty(projectID))
            {
                try
                {
                    // Project folder
                    settingsModel.ProjectRootFolder = projectID.ProjectFullPath(settingsModel.RootDrive); // J:\12\1234

                    // J:\12\1234\ + letters
                    settingsModel.DefaultSavePath = $@"{settingsModel.ProjectRootFolder.Parent.FullName}\{Settings.Default.DefaultSavePath}".Replace(settingsModel.ProjectRootTag, projectID);

                    // SaveAsPDF files and folder
                    settingsModel.XmlSaveAsPDFFolder = $@"{settingsModel.ProjectRootFolder}{Settings.Default.xmlSaveAsPDFFolder}"; // J:\12\1234\.SaveAsPDF
                    settingsModel.XmlEmployeesFile = $@"{settingsModel.XmlSaveAsPDFFolder}{Settings.Default.xmlEmployeesFile}"; // J:\12\1234\.SaveAsPDF\.SaveAsPDF_Employees.xml
                    settingsModel.XmlProjectFile = $@"{settingsModel.XmlSaveAsPDFFolder}{Settings.Default.xmlProjectFile}"; // J:\12\1234\.SaveAsPDF\.SaveAsPDF_Project.xml

                    FileFoldersHelper.CreateHiddenDirectory(settingsModel.XmlSaveAsPDFFolder);
                }
                catch (Exception ex)
                {
                    // Log the exception (you can replace this with your logging mechanism)
                    MessageBox.Show($"Error loading project settings for project ID {projectID}:\n{ex.Message}");
                }
            }

            return settingsModel;
        }

    }
}
