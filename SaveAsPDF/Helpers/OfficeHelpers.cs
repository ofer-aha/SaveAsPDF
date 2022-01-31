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
using SaveAsPDF.Properties;
using System.Web.DynamicData;
using System.Data;

namespace SaveAsPDF.Helpers
{
    public static class OfficeHelpers
    {

     
        /// <summary>
        /// Create an attachments list 
        /// </summary>
        /// <param name="email"></param>
        /// <returns></returns>
        public static List<AttachmentsModel> AttachmetsToModel(this MailItem email)
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
                
                if (attachment.Size >= Settings.Default.minAttachmentSize)
                {
                    output.Add(att);
                }
               
            }
            return output;
        }
        
        public static List<EmployeeModel> DgvEmployessToModel (this DataGridView dgv )
        {
            List<EmployeeModel> output = new List<EmployeeModel>();

            foreach (DataGridViewRow row in dgv.Rows)
            {
                EmployeeModel e = new EmployeeModel
                {
                    Id = (int)row.Index,
                    FirstName = (string)row.Cells[1].Value,
                    LastName = (string)row.Cells[2].Value,
                    EmailAddress = (string)row.Cells[3].Value
                };
                output.Add(e);
            }
            return output; 
        }
        
        /// <summary>
        /// convert a MailItem object to .MHT file ready to be saved as PDF 
        /// </summary>
        /// <param name="mailItem"></param>
        /// <param name="path"></param>
        public static void SaveToPDF (this MailItem mailItem, string path)
        {
            //cretae temp file name with uniq time stamp 
            string timeStamp = DateTime.Now.ToString("yyyyMMddHHmmssffff");
            string tFilename = $@"{Path.GetTempPath()}{timeStamp}.mht";

            
            mailItem.SaveAs(tFilename, OlSaveAsType.olMHTML);

            word.Application oWord = new word.Application();
            word.Document oDOC = oWord.Documents.Open(@tFilename, false);
            
            
            oDOC.ConvertToPDF($@"{path}\\{timeStamp}_{mailItem.Subject.SafeFileName()}.pdf");

            oDOC.Close();
            
            

        }
        /// <summary>
        /// Converts the .MHT file now saved to the user's temp directory 
        /// to PDF format and saves it to the users choosen path 
        /// </summary>
        /// <param name="oDOC"></param>
        /// <param name="oFileName"></param>
        private static void ConvertToPDF( this Document oDOC, string oFileName)
        {
            object misValue = System.Reflection.Missing.Value;
            //oDOC.ExportAsFixedFormat(@oFileName, //file name 
            //                word.WdExportFormat.wdExportFormatPDF,//export format
            //                false, //OpenAfterExport
            //                word.WdExportOptimizeFor.wdExportOptimizeForPrint, //OptimizeFor
            //                word.WdExportRange.wdExportAllDocument,  //Range
            //                (int)misValue, //From
            //                (int)misValue, //To 
            //                WdExportItem.wdExportDocumentWithMarkup, //Item
            //                false,//IncludeDocProps
            //                false, //KeepIRM
            //                WdExportCreateBookmarks.wdExportCreateWordBookmarks, //CreateBookmarks
            //                true, //DocStructureTags
            //                true, //BitmapMissingFonts
            //                false, //UseISO19005_1
            //                misValue);//FixedFormatExtClassPtr

            oDOC.SaveAs2(@oFileName, //file name 
                            word.WdExportFormat.wdExportFormatPDF);//export format


            

        }
        /// <summary>
        /// Convert List of Attchments to List of String 
        /// </summary>
        /// <param name="attachments"></param>
        /// <returns></returns>
        public static List<string> AttachmentsToString(this List<Attachment> attachments)
        {
            List<string> output = new List<string>();

            foreach (Attachment Att in attachments)
            {
                output.Add($",{Att.FileName},{Att.Size.BytesToString()}");
            }

            return output;  
        }

           

            /// <summary>
            /// extantion metud to return the attachments in mailItem to List<string> 
            /// </summary>
            /// <param name="email"></param>
            /// <returns> List<string> </returns>
            public static List<string> AttachmentsToString(this MailItem email)
        {
            Attachments mailAttachments = email.Attachments;
            List<string> output = new List<string>();

            if (mailAttachments != null)
            {

                foreach (Attachment Att in email.Attachments)
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
        /// Method to get all attachments that are NOT inline attachments (like images and stuff).
        /// </summary>
        /// <param name="mailItem">
        /// The mail item.
        /// </param>
        /// <returns>
        /// The <see cref="List"/>.
        /// </returns>
        public static List<Attachment> GetMailAttachments(this MailItem mailItem)
        {
            const string PR_ATTACH_METHOD = "http://schemas.microsoft.com/mapi/proptag/0x37050003";
            const string PR_ATTACH_FLAGS = "http://schemas.microsoft.com/mapi/proptag/0x37140003";

            var attachments = new List<Attachment>();

            // if this is a plain text email, every attachment is a non-inline attachment
            if (mailItem.BodyFormat == OlBodyFormat.olFormatPlain && mailItem.Attachments.Count > 0)
            {
                attachments.AddRange(
                    mailItem.Attachments.Cast<object>().Select(attachment => attachment as Attachment));
                return attachments;
            }

            // if the body format is RTF ...
            if (mailItem.BodyFormat == OlBodyFormat.olFormatRichText)
            {
                // add every attachment where the PR_ATTACH_METHOD property is NOT 6 (ATTACH_OLE)
                attachments.AddRange(
                    mailItem.Attachments.Cast<object>().Select(attachment => attachment as Attachment).Where(thisAttachment => (int)thisAttachment.PropertyAccessor.GetProperty(PR_ATTACH_METHOD) != 6));
            }

            // if the body format is HTML ...
            if (mailItem.BodyFormat == OlBodyFormat.olFormatHTML)
            {
                // add every attachment where the ATT_MHTML_REF property is NOT 4 (ATT_MHTML_REF)
                attachments.AddRange(
                    mailItem.Attachments.Cast<object>().Select(attachment => attachment as Attachment).Where(thisAttachment => (int)thisAttachment.PropertyAccessor.GetProperty(PR_ATTACH_FLAGS) != 4));
            }

            return attachments;
        }


        /// <summary>
        /// Save the attchment files to the choosen path. 
        /// if overWrite=True existing files will be over writen else "(x)" will be appended to file names
        /// </summary>
        /// <param name="mi"></param>
        /// <param name="path"></param>
        /// <param name="overWrite">if overWrite=True existing files will be over writen</param>
        public static void SaveAttchments(this MailItem mi, string path, bool overWrite)
        {
            var attachments = mi.Attachments;
            
            if (attachments.Count != 0)
            {
                for (int i = 1; i <= mi.Attachments.Count; i++)
                {
                    String file = Path.Combine(path, mi.Attachments[i].DisplayName);
                    
                    if (File.Exists(file) && !overWrite)
                    {
                        int x = 2;
                        
                        string[] tFile = file.Split('.');
                        file = "";
                        for (int j = 0; j< tFile.Length-1;j++)
                        {
                            file += tFile[j];   
                        }
                        while (File.Exists($@"{file}_({x}).{tFile[tFile.Length-1]}"))
                        {
                            x++;
                        }
                        mi.Attachments[i].SaveAsFile($@"{file}_({x}).{tFile[tFile.Length-1]}");
                    }
                    else
                    {
                        mi.Attachments[i].SaveAsFile(Path.Combine(path, mi.Attachments[i].DisplayName));
                    }
                    
                }
            }
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
