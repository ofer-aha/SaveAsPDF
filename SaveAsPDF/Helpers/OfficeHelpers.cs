


// Ignore Spelling: Dgv

using Microsoft.Office.Interop.Outlook;
using Microsoft.Office.Interop.Word;
using SaveAsPDF.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;
using word = Microsoft.Office.Interop.Word;

namespace SaveAsPDF.Helpers
{
    public static class OfficeHelpers
    {
        /// <summary>
        /// Converts the DataGridView to a list of EmployeeModel objects.
        /// </summary>
        /// <param name="dgv">The DataGridView to convert.</param>
        /// <returns>A list of EmployeeModel objects.</returns>
        public static List<EmployeeModel> DgvEmployeesToModel(this DataGridView dgv)
        {
            List<EmployeeModel> output = new List<EmployeeModel>();
            foreach (DataGridViewRow row in dgv.Rows)
            {
                EmployeeModel e = new EmployeeModel
                {
                    Id = row.Index,
                    FirstName = row.Cells[1].Value?.ToString(),
                    LastName = row.Cells[2].Value?.ToString(),
                    EmailAddress = row.Cells[3].Value?.ToString()
                };
                output.Add(e);
            }
            return output;
        }

        /// <summary>
        /// Convert DataGridView to String 
        /// </summary>
        /// <param name="dgv"></param>
        /// <returns></returns>
        public static string dgvEmployeesToString(this DataGridView dgv)
        {
            string output = "<tr style=\"text-align:center\"><th colspan=\"3\">מתכנן אחראי</th></tr>" +
                            "<tr style=\"text-align:center\"><th>אימייל</th><th>שם משפחה</th><th>שם פרטי</th></tr>";

            foreach (DataGridViewRow row in dgv.Rows)
            {
                output += $"<tr><td style=\"text-align:left\">{row.Cells[3].Value?.ToString()}</td>" +
                        $"<td style=\"text-align:right\">{row.Cells[2].Value?.ToString()}</td>" +
                        $"<td style=\"text-align:right\">{row.Cells[1].Value?.ToString()}</td></tr>";
            }
            return output;
        }

        /// <summary>
        /// created by AI
        /// Converts the .MHT file now saved to the user's temp directory 
        /// to PDF format and saves it to the users chosen path 
        /// </summary>
        /// <param name="mailItem"></param>
        /// <param name="outputPath"></param>
        public static void SaveToPDF(this Outlook.MailItem mailItem, string outputPath)
        {
            string timeStamp = DateTime.Now.ToString("yyyyMMddHHmmss") + "_";
            // Create a temporary MHT file
            string tempFilePath = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid()}.mht");
            mailItem.SaveAs(tempFilePath, Outlook.OlSaveAsType.olMHTML);

            try
            {
                // Open the MHT file in Word
                var wordApp = new word.Application();
                var doc = wordApp.Documents.Open(tempFilePath, ReadOnly: false);

                // Convert to PDF
                string pdfFileName = $"{Path.GetFileNameWithoutExtension(timeStamp + mailItem.Subject.SafeFolderName())}.pdf";
                string pdfPath = Path.Combine(outputPath, pdfFileName);
                doc.ExportAsFixedFormat(pdfPath, WdExportFormat.wdExportFormatPDF);

                // Clean up
                doc.Close(SaveChanges: false);
                wordApp.Quit();
            }
            catch (System.Exception ex)
            {
                MessageBox.Show($"Error converting email to PDF: {ex.Message}");
            }
            finally
            {
                // Delete the temporary MHT file
                File.Delete(tempFilePath);
            }
        }

        /// <summary>
        /// Converts the .MHT file now saved to the user's temp directory 
        /// to PDF format and saves it to the users chosen path 
        /// </summary>
        /// <param name="oDOC"></param>
        /// <param name="oFileName"></param>
        private static void ConvertToPDF(this Document oDOC, string oFileName)
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

            oDOC.SaveAs2(@oFileName,  //file name 
                            WdExportFormat.wdExportFormatPDF);//export format
        }

        /// <summary>
        /// Convert List of Attachments to List of String 
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
        /// Converts the list of attachment file names to an HTML table string.
        /// </summary>
        /// <param name="attachmentList">The list of attachment file names.</param>
        /// <param name="path">The path where the attachments are saved.</param>
        /// <returns>The HTML table string representing the attachments.</returns>
        public static string AttachmentsToString(this List<string> attachmentList, string path)
        {
            string output = string.Empty;
            if (attachmentList.Count == 0)
            {
                output = "<tr style=\"text-align:center\"><td colspan=\"3\">לא נבחרו/נמצאו קבצים מצורפים לשמירה.</td></tr>";
            }
            else
            {
                if (attachmentList.Count == 1)
                {
                    output = "<tr style=\"text-align:center\"><th colspan=\"3\">נשמר קובץ אחד</th></tr>";
                }
                else
                {
                    output = $"<tr style=\"text-align:center\"><th colspan=\"3\">נשמרו {attachmentList.Count} קבצים</th></tr>";
                }

                output += "<tr style=\"text-align:center\"><th></th><th>קובץ</th> <th>גודל</th></tr>";
                foreach (string Att in attachmentList)
                {
                    string[] t = Att.Split('|');
                    output += $"<tr style=\"text-align:left\"><td></td><td><a href='file://{Path.Combine(path, t[0])}'>{t[0]}</a></td><td>{t[1]}</td></tr>";
                }
            }

            return output;
        }

        /// <summary>
        /// Retrieves the attachments from a MailItem object.
        /// </summary>
        /// <param name="mailItem">The MailItem object.</param>
        /// <returns>A list of Attachment objects.</returns>
        public static List<Attachment> GetAttachmentsFromEmail(this MailItem mailItem)
        {
            const string PR_ATTACH_METHOD = "http://schemas.microsoft.com/mapi/proptag/0x37050003";
            const string PR_ATTACH_FLAGS = "http://schemas.microsoft.com/mapi/proptag/0x37140003";

            var attachments = new List<Attachment>();

            try
            {
                if (mailItem.BodyFormat == OlBodyFormat.olFormatPlain && mailItem.Attachments.Count > 0)
                {
                    attachments.AddRange(mailItem.Attachments.Cast<Attachment>());
                }
                else if (mailItem.BodyFormat == OlBodyFormat.olFormatRichText)
                {
                    attachments.AddRange(
                        mailItem.Attachments.Cast<Attachment>()
                            .Where(thisAttachment => (int)thisAttachment.PropertyAccessor.GetProperty(PR_ATTACH_METHOD) != 6));
                }
                else if (mailItem.BodyFormat == OlBodyFormat.olFormatHTML)
                {
                    attachments.AddRange(
                        mailItem.Attachments.Cast<Attachment>()
                            .Where(thisAttachment => (int)thisAttachment.PropertyAccessor.GetProperty(PR_ATTACH_FLAGS) != 4));
                }
            }
            catch (System.Exception ex)
            {
                // Handle exceptions (e.g., log, display a message, etc.)
                MessageBox.Show($"Error while processing attachments: {ex.Message}", "OfficeHelpers:GetAttachmentsFromEmail");
            }

            return attachments;
        }
        /// <summary>
        /// Read the selected attachments file names form the DataGridView
        /// </summary>
        /// <param name="dgv">DataGrigViwe with attachments</param>
        /// <returns>list of selected file names [list of strings]</returns>
        public static List<String> GetSelectedAttachmentFiles(this DataGridView dgv)
        {
            List<string> output = new List<string>();

            foreach (DataGridViewRow row in dgv.Rows)
            {
                if (bool.Parse(row.Cells[1].Value.ToString()))
                {
                    output.Add(row.Cells[2].Value.ToString());
                }
            }
            return output;
        }
        /// <summary>
        /// Saves the selected attachments from a MailItem object to the specified path.
        /// </summary>
        /// <param name="mi">The MailItem object.</param>
        /// <param name="dgv">The DataGridView containing the attachments.</param>
        /// <param name="path">The path where the attachments will be saved.</param>
        /// <param name="overWrite">A flag indicating whether to overwrite existing files with the same name.</param>
        /// <returns>A list of strings representing the saved attachments.</returns>
        public static List<string> SaveAttachments(this MailItem mi, DataGridView dgv, string path, bool overWrite = false)
        {
            var selectedFileNames = new HashSet<string>(dgv.GetSelectedAttachmentFiles(), StringComparer.OrdinalIgnoreCase);
            var output = new List<string>();

            foreach (Attachment attachment in mi.Attachments)
            {
                try
                {
                    string originalFileName = attachment.FileName;
                    string safeFileName = Path.GetFileNameWithoutExtension(originalFileName);
                    string extension = Path.GetExtension(originalFileName);

                    // Ensure safe filename (replace invalid characters)
                    safeFileName = string.Join("_", safeFileName.Split(Path.GetInvalidFileNameChars()));

                    // Combine the safe file name and extension using Path.ChangeExtension method
                    string targetFilePath = Path.ChangeExtension(Path.Combine(path, safeFileName), extension);

                    if (selectedFileNames.Contains(originalFileName))
                    {
                        if (File.Exists(targetFilePath) && !overWrite)
                        {
                            int suffix = 2;
                            while (File.Exists(targetFilePath))
                            {
                                targetFilePath = Path.ChangeExtension(Path.Combine(path, $"{safeFileName}_({suffix})"), extension);
                                suffix++;
                            }
                        }

                        attachment.SaveAsFile(targetFilePath);
                        output.Add($"{Path.GetFileName(targetFilePath)}|{attachment.Size.BytesToString()}");
                    }
                }
                catch (System.Exception ex)
                {
                    // Handle exceptions (e.g., invalid filenames, permissions, etc.)
                    output.Add($"Error saving attachment: {attachment.FileName} ({ex.Message})");
                }
            }

            return output;
        }

    }

}
