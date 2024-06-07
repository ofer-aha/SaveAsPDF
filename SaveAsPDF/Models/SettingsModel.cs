using System.Collections.Generic;

namespace SaveAsPDF.Models
{
    public class SettingsModel
    {
        /// <summary>
        /// List of the projectModel's folder structure
        /// </summary>
        public List<string> ProjectFolders { get; set; }
        /// <summary>
        /// The main projectModel's root folder (or drive, i.e j:\ or c:\projects) 
        /// </summary>
        public string RootFolder { get; set; }
        /// <summary>
        /// The default folder for quick mail save. 
        /// </summary>
        public string DefaultFolder { get; set; }


    }
}
