using System.IO;

namespace SaveAsPDF.Models
{
    public class SettingsModel
    {
        /// <summary>
        /// The main projectModel's root folder (or drive, i.e j:\ or c:\projects) 
        /// </summary>
        public string RootDrive { get; set; }
        /// <summary>
        ///  The folder name for all SaveAsPDF XML files 
        /// </summary>
        public string XmlSaveAsPDFFolder { get; set; }
        /// <summary>
        ///  The file name for the project data XML file 
        /// </summary>
        public string XmlProjectFile { get; set; }
        /// <summary>
        /// The file name for the employeesModel XML file 
        /// </summary>
        public string XmlEmployeesFile { get; set; }
        /// <summary>
        ///  The file name for the default tree file  
        /// </summary>
        public string DefaultTreeFile { get; set; }
        /// <summary>
        /// The default folder for quick mail save. 
        /// </summary>
        public string DefaultSavePath { get; set; }
        /// <summary>
        /// The minimum attachment size to distinguish attachment form signature image 
        /// Default recommended size in 8kb (8192b)  
        /// </summary>
        public int MinAttachmentSize { get; set; }
        /// <summary>
        /// The name-tag used to mark the current date to add to folder name (_Date_)
        /// </summary>
        public string DateTag { get; set; }
        /// <summary>
        /// The name-tag used to mark the default root folder/drive (_projrctID_)  
        /// </summary>
        public string ProjectRootTag { get; set; }
        /// <summary>
        /// The default folder index in settings form combo 
        /// </summary>
        public int DefaultFolderID { get; set; }
        /// <summary>
        /// user choose to open PDF file upon finish 
        /// </summary>
        public bool OpenPDF { get; set; }
        /// <summary>
        /// the list of the 10 last projects typed by the user
        /// it will save it as 1234;3456;6789.. and so on ';' = delimiter 
        /// </summary>
        public string LastProjects { get; set; }
        /// <summary>
        /// The amount of projects the LastProjects list will hold 
        /// default = 10 
        /// </summary>
        public int LastProjectsCount { get; set; }
        /// <summary>
        /// The project's root folder
        /// </summary>
        public DirectoryInfo ProjectRootFolders { get; set; }
    }
}
