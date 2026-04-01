using System.Collections.Generic;

namespace SaveAsPDF.Models
{
    /// <summary>
    /// Represents an Outlook contacts folder with its sub-folders.
    /// </summary>
    public class ContactFolderInfo
    {
        /// <summary>
        /// Display name of the folder or store.
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// The MAPI EntryID of the folder. Null for store-level nodes.
        /// </summary>
        public string EntryID { get; set; }

        /// <summary>
        /// The StoreID that owns this folder.
        /// </summary>
        public string StoreID { get; set; }

        /// <summary>
        /// Child contact folders.
        /// </summary>
        public List<ContactFolderInfo> SubFolders { get; set; } = new List<ContactFolderInfo>();
    }
}
