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
        /// The projectModel ID given by the user  
        /// </summary>
        public string ProjectNumber { get; set; }
        /// <summary>
        /// Project Name 
        /// </summary>
        public string ProjectName { get; set; }
        /// <summary>
        /// Send mail notification to the employee(s) responsible for the projectModel?
        /// </summary>
        public bool NoteEmployee { get; set; }
        /// <summary>
        /// The Default projectModel sub folders
        /// </summary>
        public List<string> SubFolders { get; set; }
        /// <summary>
        /// General projectModel notes left by the user
        /// </summary>
        public string ProjectNotes { get; set; }

    }
}
