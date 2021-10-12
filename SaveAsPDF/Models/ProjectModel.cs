using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SaveAsPDF.Models
{
    public class ProjectModel
    {
        /// <summary>
        /// Project ID
        /// </summary>
        public int Id { get; set; }
        /// <summary>
        /// The project ID given by the user  
        /// </summary>
        public string ProjectNumber { get; set; }
        /// <summary>
        /// Project Name 
        /// </summary>
        public string ProjectName { get; set; }
        /// <summary>
        /// The employee(s) responsible for the project
        /// </summary>
        public bool NoteEmployee { get; set; }
        /// <summary>
        /// The Default project sub folders
        /// </summary>
        public List<string> SubFolders { get; set; }


    }
}
