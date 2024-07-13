using Microsoft.Office.Interop.Outlook;
using System.Collections.Generic;

namespace SaveAsPDF.Models
{
    public class DocumentModel
    {
        /// <summary>
        /// To: field from the mailItem 
        /// </summary>
        public string ToField { get; set; }
        /// <summary>
        /// From: field from mailItem
        /// </summary>
        public string FromField { get; set; }
        /// <summary>
        /// CC: field from mailItem
        /// </summary>
        public List<string> CcField { get; set; } = new List<string>();
        /// <summary>
        /// Subject: field from mailItem
        /// </summary>
        public string Subject { get; set; }
        /// <summary>
        /// Message body
        /// </summary>
        public string Body { get; set; }
        /// <summary>
        /// Additional notes/comments added by the user while saving to PDF
        /// </summary>
        public string Notes { get; set; }
        /// <summary>
        /// List of Attachment objects 
        /// </summary>
        public List<Attachment> Attachments { get; set; } = new List<Attachment>();


    }
}
