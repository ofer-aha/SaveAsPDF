using SaveAsPDF.Properties;

namespace SaveAsPDF.Helpers
{
    public static class SettingsHelpers
    {
        public static void loadSettingsToModel()
        {
            Models.SettingsModel settingsModel = new Models.SettingsModel();

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
            settingsModel.ProjectRootFolders = Settings.Default.ProjectRootFolders;

        }


    }
}
