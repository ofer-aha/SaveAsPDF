using SaveAsPDF.Models;
using SaveAsPDF.Properties;
using System;
using System.IO;
using System.Text.RegularExpressions;

namespace SaveAsPDF.Helpers
{
    /// <summary>
    /// Provides helper methods for loading, saving, and resetting application settings.
    /// </summary>
    [Serializable]
    public static class SettingsHelpers
    {
        /// <summary>
        /// The default root drive for projects.
        /// </summary>
        static public readonly string rootDrive = @"C:\Projects\";
        /// <summary>
        /// The default folder name for SaveAsPDF XML files.
        /// </summary>
        static public readonly string xmlSaveAsPDFFolder = @".SaveAsPDF\";
        /// <summary>
        /// The default file name for the project data XML file.
        /// </summary>
        static public readonly string xmlProjectFile = @".SaveAsPDF_Project.xml";
        /// <summary>
        /// The default file name for the employees XML file.
        /// </summary>
        static public readonly string xmlEmployeesFile = @".SaveAsPDF_Emploeeys.xml";
        /// <summary>
        /// The default file name for the tree file.
        /// </summary>
        static public readonly string defaultTreeFile = @"C:\Projects\tree.fld";
        /// <summary>
        /// The minimum attachment size in bytes (default: 8192).
        /// </summary>
        static public readonly int minAttachmentSize = 8192;
        /// <summary>
        /// The tag used to mark the current date in folder names.
        /// </summary>
        static public readonly string dateTag = "_תאריך_";
        /// <summary>
        /// The default folder ID.
        /// </summary>
        static public readonly int defaultFolderID = 1;
        /// <summary>
        /// The tag used to mark the project root in folder names.
        /// </summary>
        static public readonly string projectRootTag = "_מספר_פרויקט_";
        /// <summary>
        /// Indicates whether to open the PDF after saving.
        /// </summary>
        static public readonly bool openPDF = false;
        /// <summary>
        /// The default last projects string.
        /// </summary>
        static public readonly string lastProjects = "1000;";
        /// <summary>
        /// The default number of last projects to keep.
        /// </summary>
        static public readonly int lastProjectsCount = 10;
        /// <summary>
        /// The default project root folders path.
        /// </summary>
        static public readonly string sProjectRootFolders = $@"{rootDrive}10\1000\";
        /// <summary>
        /// The default save path for files.
        /// </summary>
        static public readonly string defaultSavePath = $@"{rootDrive}{sProjectRootFolders}Inbox\";

        /// <summary>
        /// Loads the application settings from <see cref="Settings.Default"/> into the provided <see cref="SettingsModel"/>.
        /// </summary>
        /// <param name="settingsModel">The <see cref="SettingsModel"/> to populate with settings.</param>
        /// <returns>The populated <see cref="SettingsModel"/>.</returns>
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
                XMessageBox.Show(
                    e.Message,
                    "שגיאה בטעינת הגדרות",
                    XMessageBoxButtons.OK,
                    XMessageBoxIcon.Error,
                    XMessageAlignment.Right,
                    XMessageLanguage.Hebrew
                );
            }

            return settingsModel;
        }

        /// <summary>
        /// Saves the provided <see cref="SettingsModel"/> to <see cref="Settings.Default"/>.
        /// </summary>
        /// <param name="settingsModel">The <see cref="SettingsModel"/> to save.</param>
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
                XMessageBox.Show(
                    e.Message,
                    "שגיאה בשמירת הגדרות",
                    XMessageBoxButtons.OK,
                    XMessageBoxIcon.Error,
                    XMessageAlignment.Right,
                    XMessageLanguage.Hebrew
                );
            }
        }

        /// <summary>
        /// Loads the default settings into <see cref="Settings.Default"/> for first run or reset.
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
                XMessageBox.Show(
                    e.Message,
                    "שגיאה בטעינת הגדרות ברירת מחדל",
                    XMessageBoxButtons.OK,
                    XMessageBoxIcon.Error,
                    XMessageAlignment.Right,
                    XMessageLanguage.Hebrew
                );
            }
        }

        /// <summary>
        /// Loads the settings model from <see cref="Settings.Default"/>. Parameter-less overload.
        /// </summary>
        /// <returns>The loaded <see cref="SettingsModel"/>.</returns>
        public static SettingsModel LoadProjectSettings()
        {
            return LoadProjectSettings(string.Empty);
        }

        /// <summary>
        /// Loads the settings model from <see cref="Settings.Default"/> for a specific project ID.
        /// </summary>
        /// <param name="projectID">The project ID to use for loading project-specific settings.</param>
        /// <returns>The loaded <see cref="SettingsModel"/>.</returns>
        public static SettingsModel LoadProjectSettings(this string projectID)
        {
            try
            {
                // Load the root drive
                FormMain.settingsModel.RootDrive = Settings.Default.RootDrive ?? rootDrive;

                // Folder tags
                FormMain.settingsModel.DateTag = Settings.Default.DateTag ?? dateTag;
                FormMain.settingsModel.ProjectRootTag = Settings.Default.ProjectRootTag ?? projectRootTag;

                // More settings
                FormMain.settingsModel.DefaultFolderID = Settings.Default.DefaultFolderID > 0 ? Settings.Default.DefaultFolderID : defaultFolderID;
                FormMain.settingsModel.MinAttachmentSize = Settings.Default.MinAttachmentSize > 0 ? Settings.Default.MinAttachmentSize : minAttachmentSize;
                FormMain.settingsModel.DefaultTreeFile = Settings.Default.DefaultTreeFile ?? defaultTreeFile;
                FormMain.settingsModel.OpenPDF = Settings.Default.OpenPDF;

                if (!string.IsNullOrEmpty(projectID))
                {
                    // Project folder
                    FormMain.settingsModel.ProjectRootFolder = projectID.ProjectFullPath(FormMain.settingsModel.RootDrive);
                    string projectRootPath = FormMain.settingsModel.ProjectRootFolder.FullName;

                    // Default save path based on the project root and settings selection
                    string defaultSaveSetting = Settings.Default.DefaultSavePath ?? string.Empty;
                    defaultSaveSetting = defaultSaveSetting.Replace(FormMain.settingsModel.ProjectRootTag, projectID);
                    defaultSaveSetting = defaultSaveSetting.Replace(FormMain.settingsModel.DateTag, DateTime.Now.ToString("dd.MM.yyyy"));

                    FormMain.settingsModel.DefaultSavePath = ToAbsolutePath(projectRootPath, defaultSaveSetting, fallbackToBase: true);

                    // set the SaveAsPDF files and folder paths
                    string xmlFolderSetting = Settings.Default.xmlSaveAsPDFFolder ?? xmlSaveAsPDFFolder;
                    FormMain.settingsModel.XmlSaveAsPDFFolder = ToAbsolutePath(projectRootPath, xmlFolderSetting, fallbackToBase: false);

                    string xmlEmployeesSetting = Settings.Default.xmlEmployeesFile ?? xmlEmployeesFile;
                    string xmlProjectSetting = Settings.Default.xmlProjectFile ?? xmlProjectFile;

                    FormMain.settingsModel.XmlEmployeesFile = ToAbsolutePath(FormMain.settingsModel.XmlSaveAsPDFFolder, xmlEmployeesSetting, fallbackToBase: false);
                    FormMain.settingsModel.XmlProjectFile = ToAbsolutePath(FormMain.settingsModel.XmlSaveAsPDFFolder, xmlProjectSetting, fallbackToBase: false);

                    FileFoldersHelper.CreateHiddenDirectory(FormMain.settingsModel.XmlSaveAsPDFFolder);
                }
            }
            catch (Exception ex)
            {
                XMessageBox.Show(
                    $"שגיאה בטעינת הגדרות פרויקט עבור מזהה {projectID}:\n{ex.Message}",
                    "שגיאה בטעינת הגדרות פרויקט",
                    XMessageBoxButtons.OK,
                    XMessageBoxIcon.Error,
                    XMessageAlignment.Right,
                    XMessageLanguage.Hebrew
                );
            }

            return FormMain.settingsModel;
        }

        private static string ToAbsolutePath(string basePath, string value, bool fallbackToBase)
        {
            string normalizedBase = NormalizePathToken(basePath);
            string normalizedValue = NormalizePathToken(value);

            if (string.IsNullOrWhiteSpace(normalizedValue))
                return fallbackToBase ? normalizedBase : Path.Combine(normalizedBase, string.Empty);

            try
            {
                if (Path.IsPathRooted(normalizedValue))
                    return Path.GetFullPath(normalizedValue);

                return Path.GetFullPath(Path.Combine(normalizedBase, normalizedValue.TrimStart('\\', '/')));
            }
            catch
            {
                if (fallbackToBase)
                    return normalizedBase;

                return Path.Combine(normalizedBase, normalizedValue.TrimStart('\\', '/'));
            }
        }

        private static string NormalizePathToken(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
                return string.Empty;

            string normalized = value.Trim();
            normalized = PathBreadcrumbHelper.ToPathFromBreadcrumb(normalized);
            normalized = Regex.Replace(normalized, @"\s*([\\/])\s*", "$1");
            normalized = normalized.Replace('/', '\\');

            if (normalized.Length == 2 && normalized[1] == ':')
                normalized += "\\";

            return normalized;
        }
    }
}
