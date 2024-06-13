namespace SaveAsPDF.Models
{
    public class AttachmentsModel

    {
        /// <summary>
        /// File ID in the list.... future use? 
        /// </summary>
        public int attachmentId { get; set; }
        /// <summary>
        /// The checked-box in the dataGridView. Default: True
        /// </summary>
        public bool isChecked { get; set; } = true;
        /// <summary>
        /// File name... nothing to say about this 
        /// </summary>
        public string fileName { get; set; }
        /// <summary>
        /// File size.... see file name... 
        /// </summary>
        public string fileSize { get; set; }

    }

}
