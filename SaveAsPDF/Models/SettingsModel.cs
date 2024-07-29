using System;
using System.IO;

namespace SaveAsPDF.Models
{

    [Serializable]
    public class SettingsModel
    {
        /// <summary>
        /// The main _projectModel's root folder (or drive, i.e j:\ or c:\projects) 
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
        /// The file name for the _employeesModel XML file 
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
        /// The project's root folder
        /// </summary>
        public DirectoryInfo ProjectRootFolder { get; set; }

        public SettingsModel()
        {
            //reset fields to default values
            RootDrive = @"C:\Projects\";
            XmlSaveAsPDFFolder = @".SaveAsPDF\";
            XmlProjectFile = @".SaveAsPDF_Project.xml";
            XmlEmployeesFile = @".SaveAsPDF_Emploeeys.xml";
            DefaultTreeFile = @"C:\Projects\tree.fld";
            MinAttachmentSize = 8192;
            DateTag = "_תאריך_";
            DefaultFolderID = 1;
            ProjectRootTag = "_מספר_פרויקט_";
            OpenPDF = false;
            ProjectRootFolder = new DirectoryInfo($@"{RootDrive}10\1000\");
            DefaultSavePath = $@"{RootDrive}{ProjectRootFolder}Inbox\";
        }



    }
}
