using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection.Emit;
using System.Text;
using System.Threading.Tasks;

namespace SaveAsPDF.Models
{
    public class SettingsModel
    {
        /// <summary>
        /// List of the project's folder structure
        /// </summary>
        public List<string> ProjectFolders { get; set; }
        /// <summary>
        /// The main project's root folder (or drive, i.e j:\ or c:\projects) 
        /// </summary>
        public string RootFolder { get; set; }
        /// <summary>
        /// The default folder for quick mail save. 
        /// </summary>
        public string DefaultFolder { get; set; }

        
    }
}
