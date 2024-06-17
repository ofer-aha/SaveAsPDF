using System.Collections.Generic;

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
        public List<string> SubFolders { get; set; }
        /// <summary>
        /// General _projectModel notes left by the user
        /// </summary>
        public string ProjectNotes { get; set; }

    }
}
