using Microsoft.Office.Interop.Outlook;
using Microsoft.Office.Interop.Word;
using Outlook = Microsoft.Office.Interop.Outlook;
using word = Microsoft.Office.Interop.Word;
using SaveAsPDF.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Globalization;
using System.IO;
using System.Windows.Forms;

namespace SaveAsPDF.Helpers
{
    public static class OfficeHelpers
    {

     

        public static List<AttachmentsModel> AttachmetsToModel(this Outlook.MailItem email)
        {
            List<AttachmentsModel> output = new List<AttachmentsModel>();
            int i = 0; 
            foreach (Outlook.Attachment attachment in email.Attachments)
            {
                i += 1;
                AttachmentsModel att = new AttachmentsModel
                {
                    attachmentId = i,
                    isChecked = true,
                    fileName = attachment.FileName,
                    fileSize = attachment.Size.BytesToString()

                };


                output.Add(att);
            }
            return output;
        }
        
        public static List<EmployeeModel> DgvEmployessToModel (this DataGridView dgv )
        {
            EmployeeModel e = new EmployeeModel();
            List<EmployeeModel> output = new List<EmployeeModel>();
            
            foreach (DataGridViewRow row  in dgv.Rows)
            {
                int t = 0;
                if( int.TryParse(row.Cells[0].ToString(), NumberStyles.Integer, null, out t))
                {
                    e.Id = t;
                }
                e.FirstName = row.Cells[1].ToString();
                e.LastName = row.Cells[2].ToString();
                e.EmailAddress = row.Cells[3].ToString();
                output.Add(e);
            }
            return output; 
        }
        
        
        /// <summary>
        /// extantion metud to return the attachments in mailItem to List<string> 
        /// </summary>
        /// <param name="email"></param>
        /// <returns> List<string> </returns>
        public static List<string> AttachmentsToString(this Outlook.MailItem email)
        {
            
            Outlook.Attachments mailAttachments = email.Attachments;
            List<string> output = new List<string>();

            if (mailAttachments != null)
            {
            
                foreach (Outlook.Attachment Att in email.Attachments)
                {
                    output.Add($"{ Att.FileName},{Att.Size.BytesToString()}");
                }

                //    for (int i = 1; i <= mailAttachments.Count; i++)
                //{
                //    Outlook.Attachment currentAttachment = mailAttachments[i];
                //    if (currentAttachment != null)
                //    {
                //        //attachmentInfo.AppendFormat("#{0}\n\rFile name: {1}\n\rFile Size: {2}\n\rType: {3}\n\n\r",
                //        //                                i, currentAttachment.FileName, currentAttachment.Size, currentAttachment.Type);
                //        //return new string[] { currentAttachment.FileName, currentAttachment.Size.BytesToString() }));

                //        //Marshal.ReleaseComObject(currentAttachment);
                //    }
                //}
                //}
                ////    if (attachmentInfo.Length > 0)
                //        toolStripStatusLabel1.Text = "E-mail attachments: " + attachmentInfo.ToString();
                //Marshal.ReleaseComObject(mailAttachments);
            }

            return output;
        }
        /// <summary>
        /// convert a MailItem object to .MHT file ready to be saved as PDF 
        /// </summary>
        /// <param name="mailItem"></param>
        /// <param name="path"></param>
        public static void SaveToPDF (Outlook.MailItem mailItem, string path)
        {
            //   cretae uniq timestamp for uniqe fimenames
            string TimeStamp = DateTime.Now.ToString("yyyyMMddHHmmssffff", CultureInfo.CurrentUICulture);
            string tFilename = $@"{Path.GetTempPath()}{TimeStamp}.mht";

            mailItem.SaveAs(tFilename, OlSaveAsType.olMHTML);

            word.Application oWord = new word.Application();
            word.Document oDOC = oWord.Documents.Open(@tFilename, true);
            object misValue = System.Reflection.Missing.Value;

            string oFileName = $@"{ path }{ TimeStamp }_{ mailItem.Subject.SafeFileName() }.pdf";

            ConvertPDF(oDOC, misValue, @oFileName);

        }
        /// <summary>
        /// Converts the .MHT file now saved to the user's temp directory 
        /// to PDF format and saves it to the users choosen path 
        /// </summary>
        /// <param name="oDOC"></param>
        /// <param name="misValue"></param>
        /// <param name="oFileName"></param>
        private static void ConvertPDF(Document oDOC, object misValue, string oFileName)
        {
            oDOC.ExportAsFixedFormat(@oFileName,
                            word.WdExportFormat.wdExportFormatPDF,
                            false,
                            word.WdExportOptimizeFor.wdExportOptimizeForPrint,
                            word.WdExportRange.wdExportAllDocument, (int)misValue, (int)misValue, WdExportItem.wdExportDocumentWithMarkup, false,
                             true, WdExportCreateBookmarks.wdExportCreateWordBookmarks, true, true, false, misValue);
        }


        //   cretae uniq timestamp for uniqe fimenames

        //    string TimeStamp = DateTime.Now.ToString("yyyyMMddHHmmssffff", CultureInfo.CurrentUICulture);

        //    string tFilename = @Path.GetTempPath() + TimeStamp + ".mht";

        //    mailItem.SaveAs(@tFilename, OlSaveAsType.olMHTML);

        //        word.Application oWord = new word.Application();
        //    word.Document oDOC = oWord.Documents.Open(@tFilename, true);

        //    object misValue = System.Reflection.Missing.Value;

        //    string oFileName = $"{ lblFolder.Text }{ TimeStamp }_{ mailItem.Subject.SafeFileName() }.pdf";

        //    ConvertPDF(oDOC, misValue, @oFileName);
        //}

        //private static void ConvertPDF(Document oDOC, object misValue, string oFileName)
        //{
        //    oDOC.ExportAsFixedFormat(@oFileName,
        //                    word.WdExportFormat.wdExportFormatPDF,
        //                    false,
        //                    word.WdExportOptimizeFor.wdExportOptimizeForPrint,
        //                    word.WdExportRange.wdExportAllDocument, (int)misValue, (int)misValue, WdExportItem.wdExportDocumentWithMarkup, false,
        //                     true, WdExportCreateBookmarks.wdExportCreateWordBookmarks, true, true, false, misValue);
        //}





    }

}
