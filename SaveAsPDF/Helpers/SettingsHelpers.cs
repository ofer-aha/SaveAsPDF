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
        // Readonly default settings
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
                // Helper method to get a setting or its default value
                T GetSettingOrDefault<T>(T settingValue, T defaultValue)
                {
                    return settingValue != null ? settingValue : defaultValue;
                }

                // Load the settings from Settings.Default
                settingsModel.RootDrive = GetSettingOrDefault(Settings.Default.RootDrive, rootDrive);
                settingsModel.XmlSaveAsPDFFolder = GetSettingOrDefault(Settings.Default.xmlSaveAsPDFFolder, xmlSaveAsPDFFolder);
                settingsModel.XmlProjectFile = GetSettingOrDefault(Settings.Default.xmlProjectFile, xmlProjectFile);
                settingsModel.XmlEmployeesFile = GetSettingOrDefault(Settings.Default.xmlEmployeesFile, xmlEmployeesFile);
                settingsModel.DefaultTreeFile = GetSettingOrDefault(Settings.Default.DefaultTreeFile, defaultTreeFile);
                settingsModel.DefaultSavePath = GetSettingOrDefault(Settings.Default.DefaultSavePath, defaultSavePath);
                settingsModel.MinAttachmentSize = Settings.Default.MinAttachmentSize > 0 ? Settings.Default.MinAttachmentSize : minAttachmentSize;
                settingsModel.DateTag = GetSettingOrDefault(Settings.Default.DateTag, dateTag);
                settingsModel.DefaultFolderID = Settings.Default.DefaultFolderID > 0 ? Settings.Default.DefaultFolderID : defaultFolderID;
                settingsModel.ProjectRootTag = GetSettingOrDefault(Settings.Default.ProjectRootTag, projectRootTag);
                settingsModel.OpenPDF = Settings.Default.OpenPDF;

                // Set the ProjectRootFolder
                settingsModel.ProjectRootFolder = !string.IsNullOrEmpty(Settings.Default.sProjectRootFolders)
                    ? new DirectoryInfo(Settings.Default.sProjectRootFolders)
                    : new DirectoryInfo(rootDrive);
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "SettingsHelpers:LoadSettingsToModel");
            }

            return settingsModel;
        }

        /// <summary>
        /// Save the _settingsModel to Settings.Settings 
        /// </summary>
        /// <param name="settingsModel"><see cref="SettingsModel"/> object</param>
        public static void SaveModelToSettings(SettingsModel settingsModel)
        {
            try
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
                Settings.Default.sProjectRootFolders = settingsModel.ProjectRootFolder?.FullName ?? rootDrive;
                Settings.Default.Save();
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "SettingsHelpers:SaveModelToSettings");
            }
        }

        /// <summary>
        /// Load default settings for first run or reset
        /// </summary>
        public static void LoadDefaultSettings()
        {
            try
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
                Settings.Default.sProjectRootFolders = sProjectRootFolders;
                Settings.Default.Save();
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "SettingsHelpers:LoadDefaultSettings");
            }
        }

        /// <summary>
        /// Load the settingsModel from Settings.Settings. Parameterless overload.  
        /// </summary>
        /// <returns><see cref="SettingsModel"/> object </returns>
        public static SettingsModel LoadProjectSettings()
        {
            return LoadProjectSettings(string.Empty);
        }

        /// <summary>
        /// Load the settingsModel from Settings.Settings 
        /// </summary>
        /// <param name="projectID"><see cref="string"/> projectID</param>
        /// <returns><see cref="SettingsModel"/> object </returns>
        public static SettingsModel LoadProjectSettings(this string projectID)
        {
            SettingsModel settingsModel = new SettingsModel();

            try
            {
                settingsModel.RootDrive = Settings.Default.RootDrive ?? rootDrive;

                // Folder tags
                settingsModel.DateTag = Settings.Default.DateTag ?? dateTag;
                settingsModel.ProjectRootTag = Settings.Default.ProjectRootTag ?? projectRootTag;

                // More settings
                settingsModel.DefaultFolderID = Settings.Default.DefaultFolderID > 0 ? Settings.Default.DefaultFolderID : defaultFolderID;
                settingsModel.MinAttachmentSize = Settings.Default.MinAttachmentSize > 0 ? Settings.Default.MinAttachmentSize : minAttachmentSize;
                settingsModel.DefaultTreeFile = Settings.Default.DefaultTreeFile ?? defaultTreeFile;
                settingsModel.OpenPDF = Settings.Default.OpenPDF;

                if (!string.IsNullOrEmpty(projectID))
                {
                    // Project folder
                    settingsModel.ProjectRootFolder = projectID.ProjectFullPath(settingsModel.RootDrive);

                    // Default save path
                    settingsModel.DefaultSavePath = $@"{settingsModel.ProjectRootFolder.Parent.FullName}\{Settings.Default.DefaultSavePath}".Replace(settingsModel.ProjectRootTag, projectID);

                    // SaveAsPDF files and folder
                    settingsModel.XmlSaveAsPDFFolder = $@"{settingsModel.ProjectRootFolder}{Settings.Default.xmlSaveAsPDFFolder}";
                    settingsModel.XmlEmployeesFile = $@"{settingsModel.XmlSaveAsPDFFolder}{Settings.Default.xmlEmployeesFile}";
                    settingsModel.XmlProjectFile = $@"{settingsModel.XmlSaveAsPDFFolder}{Settings.Default.xmlProjectFile}";

                    FileFoldersHelper.CreateHiddenDirectory(settingsModel.XmlSaveAsPDFFolder);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error loading project settings for project ID {projectID}:\n{ex.Message}", "SettingsHelpers:LoadProjectSettings");
            }

            return settingsModel;
        }
    }
}
