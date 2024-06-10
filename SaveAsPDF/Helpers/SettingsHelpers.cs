using SaveAsPDF.Properties;

namespace SaveAsPDF.Helpers
{
    public static class SettingsHelpers
    {
        /// <summary>
        /// Load the settings.settings to the settingsModel 
        /// </summary>
        public static void loadSettingsToModel()
        {
            //Models.SettingsModel settingsModel = new Models.SettingsModel();

            frmMain.settingsModel.RootDrive = Settings.Default.RootDrive;
            frmMain.settingsModel.XmlSaveAsPDFFolder = Settings.Default.xmlSaveAsPDFFolder;
            frmMain.settingsModel.XmlProjectFile = Settings.Default.xmlProjectFile;
            frmMain.settingsModel.XmlEmployeesFile = Settings.Default.xmlEmployeesFile;
            frmMain.settingsModel.DefaultTreeFile = Settings.Default.DefaultTreeFile;
            frmMain.settingsModel.DefaultSavePath = Settings.Default.DefaultSavePath;
            frmMain.settingsModel.MinAttachmentSize = Settings.Default.MinAttachmentSize;
            frmMain.settingsModel.DateTag = Settings.Default.DateTag;
            frmMain.settingsModel.DefaultFolderID = Settings.Default.DefaultFolderID;
            frmMain.settingsModel.ProjectRootTag = Settings.Default.ProjectRootTag;
            frmMain.settingsModel.OpenPDF = Settings.Default.OpenPDF;
            frmMain.settingsModel.LastProjects = Settings.Default.LastProjects;
            frmMain.settingsModel.LastProjectsCount = Settings.Default.LastProjectsCount;
            frmMain.settingsModel.ProjectRootFolders = Settings.Default.ProjectRootFolders;

        }
        /// <summary>
        /// Save the SettingsModel to Settings.Settings 
        /// </summary>
        public static void saveSettingsToModel()
        {
            Settings.Default.RootDrive = frmMain.settingsModel.RootDrive;
            Settings.Default.xmlSaveAsPDFFolder = frmMain.settingsModel.XmlSaveAsPDFFolder;
            Settings.Default.xmlProjectFile = frmMain.settingsModel.XmlProjectFile;
            Settings.Default.xmlEmployeesFile = frmMain.settingsModel.XmlEmployeesFile;
            Settings.Default.DefaultTreeFile = frmMain.settingsModel.DefaultTreeFile;
            Settings.Default.DefaultSavePath = frmMain.settingsModel.DefaultSavePath;
            Settings.Default.MinAttachmentSize = frmMain.settingsModel.MinAttachmentSize;
            Settings.Default.DateTag = frmMain.settingsModel.DateTag;
            Settings.Default.DefaultFolderID = frmMain.settingsModel.DefaultFolderID;
            Settings.Default.ProjectRootTag = frmMain.settingsModel.ProjectRootTag;
            Settings.Default.OpenPDF = frmMain.settingsModel.OpenPDF;
            Settings.Default.LastProjects = frmMain.settingsModel.LastProjects;
            Settings.Default.LastProjectsCount = frmMain.settingsModel.LastProjectsCount;
            Settings.Default.ProjectRootFolders = frmMain.settingsModel.ProjectRootFolders;
        }


    }
}
