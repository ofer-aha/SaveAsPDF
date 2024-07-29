namespace SaveAsPDF.Models
{
    public class ProjectModel
    {
        /// <summary>
        /// Project ID
        /// </summary>
        public int Id { get; set; }
        /// <summary>
        /// The _projectModel ID given by the user  
        /// </summary>
        public string ProjectNumber { get; set; }
        /// <summary>
        /// Project Name 
        /// </summary>
        public string ProjectName { get; set; }
        /// <summary>
        /// Send mail notification to the employee(s) responsible for the _projectModel?
        /// </summary>
        public bool NoteEmployee { get; set; }
        /// <summary>
        /// The Default _projectModel sub folders
        /// </summary>
        public string DefaultSaveFolder { get; set; }
        /// <summary>
        /// General _projectModel notes left by the user
        /// </summary>
        public string ProjectNotes { get; set; }

        /// <summary>
        /// The last save path of the project
        /// </summary>
        public string LastSavePath { get; set; }

        public ProjectModel()
        {
            Id = 0;
            ProjectNumber = string.Empty;
            ProjectName = string.Empty;
            NoteEmployee = false;
            DefaultSaveFolder = string.Empty;
            ProjectNotes = string.Empty;
            LastSavePath = string.Empty;
        }
    }
}
