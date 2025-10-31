using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Text;
using System.Linq;

namespace SaveAsPDF.Helpers
{
    public static class HtmlHelper
    {
        public static string GenerateHtmlContent(
            string sPath,
            List<SaveAsPDF.Models.EmployeeModel> employees,
            List<SaveAsPDF.Models.AttachmentsModel> attachments,
            string projectName,
            string projectID,
            string notes,
            string userName,
            string mailSubject = null) // optional: use email subject for pdf filename
        {
            // Prepare and precompute
            projectName = projectName ?? string.Empty;
            projectID = projectID ?? string.Empty;
            notes = notes ?? string.Empty;
            userName = userName ?? string.Empty;
            mailSubject = mailSubject ?? string.Empty;

            // Prefer mail subject for PDF filename when provided, otherwise fallback to project name
            string pdfBase = !string.IsNullOrWhiteSpace(mailSubject) ? mailSubject.SafeFolderName() : projectName.SafeFolderName();
            string pdfFileName = pdfBase + ".pdf";
            string pdfFullPath = Path.Combine(sPath ?? string.Empty, pdfFileName);
            string pdfHref = "file:///" + pdfFullPath.Replace("\\", "/");

            string encProjectName = WebUtility.HtmlEncode(projectName);
            string encProjectID = WebUtility.HtmlEncode(projectID);
            string encNotes = WebUtility.HtmlEncode(notes).Replace(Environment.NewLine, "<br>");
            string encUser = WebUtility.HtmlEncode(userName);

            // Precompute attachment display names and hrefs to avoid repeated work
            var attachmentDisplayNames = new List<string>();
            var attachmentSizes = new List<string>();
            var attachmentHrefs = new List<string>();
            if (attachments != null)
            {
                attachmentDisplayNames.Capacity = attachments.Count;
                attachmentSizes.Capacity = attachments.Count;
                attachmentHrefs.Capacity = attachments.Count;
                foreach (var att in attachments)
                {
                    var original = att?.fileName ?? string.Empty;
                    var size = att?.fileSize ?? string.Empty;

                    // Build href using the selected save path and the original filename
                    // Escape the filename for use in a file URI
                    string fileHref = "file:///" + Path.Combine(sPath ?? string.Empty, original).Replace("\\", "/");

                    attachmentDisplayNames.Add(WebUtility.HtmlEncode(original));
                    attachmentSizes.Add(WebUtility.HtmlEncode(size));
                    attachmentHrefs.Add(WebUtility.HtmlEncode(fileHref));
                }
            }

            // Precompute leader display
            string leaderDisplay = GetLeaderDisplay(employees);

            // Build into StringWriter (backed by StringBuilder) using estimated capacity
            int estimatedCapacity =2048 + (employees?.Count ??0) *120 + (attachments?.Count ??0) *160;
            var sw = new StringWriter(new StringBuilder(estimatedCapacity));
            WriteHtml(sw, encProjectName, encProjectID, encNotes, encUser, leaderDisplay, pdfFileName, pdfHref, sPath, employees, attachmentDisplayNames, attachmentSizes, attachmentHrefs);
            return sw.ToString();
        }

        // New API: write the HTML directly to a file to avoid constructing a large in-memory string
        public static void GenerateHtmlToFile(
            string outputFilePath,
            string sPath,
            List<SaveAsPDF.Models.EmployeeModel> employees,
            List<SaveAsPDF.Models.AttachmentsModel> attachments,
            string projectName,
            string projectID,
            string notes,
            string userName,
            string mailSubject = null) // optional: use email subject for pdf filename
        {
            if (string.IsNullOrEmpty(outputFilePath))
                throw new ArgumentException("outputFilePath must be provided", nameof(outputFilePath));

            // Prepare same precomputation as GenerateHtmlContent
            projectName = projectName ?? string.Empty;
            projectID = projectID ?? string.Empty;
            notes = notes ?? string.Empty;
            userName = userName ?? string.Empty;
            mailSubject = mailSubject ?? string.Empty;

            string pdfBase = !string.IsNullOrWhiteSpace(mailSubject) ? mailSubject.SafeFolderName() : projectName.SafeFolderName();
            string pdfFileName = pdfBase + ".pdf";
            string pdfFullPath = Path.Combine(sPath ?? string.Empty, pdfFileName);
            string pdfHref = "file:///" + pdfFullPath.Replace("\\", "/");

            string encProjectName = WebUtility.HtmlEncode(projectName);
            string encProjectID = WebUtility.HtmlEncode(projectID);
            string encNotes = WebUtility.HtmlEncode(notes).Replace(Environment.NewLine, "<br>");
            string encUser = WebUtility.HtmlEncode(userName);

            var attachmentDisplayNames = new List<string>();
            var attachmentSizes = new List<string>();
            var attachmentHrefs = new List<string>();
            if (attachments != null)
            {
                attachmentDisplayNames.Capacity = attachments.Count;
                attachmentSizes.Capacity = attachments.Count;
                attachmentHrefs.Capacity = attachments.Count;
                foreach (var att in attachments)
                {
                    var original = att?.fileName ?? string.Empty;
                    var size = att?.fileSize ?? string.Empty;
                    string fileHref = "file:///" + Path.Combine(sPath ?? string.Empty, original).Replace("\\", "/");

                    attachmentDisplayNames.Add(WebUtility.HtmlEncode(original));
                    attachmentSizes.Add(WebUtility.HtmlEncode(size));
                    attachmentHrefs.Add(WebUtility.HtmlEncode(fileHref));
                }
            }

            string leaderDisplay = GetLeaderDisplay(employees);

            // Write directly to file using UTF8
            using (var writer = new StreamWriter(outputFilePath, false, Encoding.UTF8))
            {
                WriteHtml(writer, encProjectName, encProjectID, encNotes, encUser, leaderDisplay, pdfFileName, pdfHref, sPath, employees, attachmentDisplayNames, attachmentSizes, attachmentHrefs);
            }
        }

        // Helper: build leader display string
        private static string GetLeaderDisplay(List<SaveAsPDF.Models.EmployeeModel> employees)
        {
            if (employees == null) return string.Empty;
            var leader = employees.FirstOrDefault(e => e.IsLeader);
            if (leader == null) return string.Empty;
            var first = WebUtility.HtmlEncode(leader.FirstName ?? string.Empty);
            var last = WebUtility.HtmlEncode(leader.LastName ?? string.Empty);
            var email = WebUtility.HtmlEncode(leader.EmailAddress ?? string.Empty);
            var name = (first + " " + last).Trim();
            return string.IsNullOrWhiteSpace(name)
                ? email
                : (name + (string.IsNullOrWhiteSpace(email) || name.Equals(email, StringComparison.OrdinalIgnoreCase) ? string.Empty : " (" + email + ")"));
        }

        // Core HTML writer used by both string and file-based outputs
        private static void WriteHtml(TextWriter writer,
            string encProjectName,
            string encProjectID,
            string encNotes,
            string encUser,
            string leaderDisplay,
            string pdfFileName,
            string pdfHref,
            string saveFolder,
            List<SaveAsPDF.Models.EmployeeModel> employees,
            List<string> attachmentDisplayNames,
            List<string> attachmentSizes,
            List<string> attachmentHrefs)
        {
            // Compose header and styles into a single chunk (tighter table styles)
            var header = new StringBuilder();
            header.AppendLine("<!doctype html>");
            header.AppendLine("<html lang=\"he\"> ");
            header.AppendLine("<head>");
            header.AppendLine("<meta charset=\"utf-8\"> ");
            header.AppendLine("<meta name=\"viewport\" content=\"width=device-width, initial-scale=1\"> ");
            header.AppendLine("<title>" + encProjectName + " - SaveAsPDF</title>");

            header.AppendLine("<style>");
            header.AppendLine(":root { font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, 'Noto Sans', Arial, sans-serif; color:#222; background:#fff; }");
            header.AppendLine("body { margin:12px; direction:rtl; background:#f6f8fa; }");
            header.AppendLine(".container { max-width:900px; margin:0 auto; }");
            header.AppendLine(".card { border:1px solid #e1e4e8; border-radius:6px; padding:8px; background:#fff; margin-bottom:10px; }");
            header.AppendLine(".header { display:flex; align-items:center; justify-content:space-between; margin-bottom:6px; }");
            header.AppendLine("h1 { font-size:16px; margin:0; }");
            header.AppendLine(".meta { font-size:12px; color:#444; }");
            header.AppendLine("table { width:100%; border-collapse:collapse; margin-top:6px; }");
            header.AppendLine("th, td { padding:6px8px; border-bottom:1px solid #f1f1f1; text-align:right; font-size:12px; }");
            header.AppendLine("th { background:#fafafa; font-weight:600; color:#333; }");
            header.AppendLine("tr:nth-child(even) td { background:#fbfbfb; }");
            header.AppendLine("a.file { color:#0366d6; text-decoration:none; }");
            header.AppendLine("a.file:hover { text-decoration:underline; }");
            header.AppendLine(".leader { font-weight:700; margin-top:4px; }");
            header.AppendLine(".note { white-space:pre-wrap; font-size:12px; color:#333; margin-top:6px; }");
            header.AppendLine(".muted { color:#666; font-size:11px; }");
            header.AppendLine("@media (max-width:600px){ body{margin:8px;} .container{padding:6px;} th, td{padding:6px;} h1{font-size:14px;} }");
            header.AppendLine("</style>");

            header.AppendLine("</head>");
            header.AppendLine("<body>");
            header.AppendLine("<div class=\"container\"> ");

            writer.Write(header.ToString());

            // Project card as a single chunk
            var projectSb = new StringBuilder();
            projectSb.AppendLine("<section class=\"card\"> ");
            projectSb.AppendLine(" <div class=\"header\"> ");
            projectSb.AppendLine(" <div>\n <h1>" + encProjectName + "</h1>\n <div class=\"meta\">מספר פרויקט: " + encProjectID + "</div>\n </div>");
            projectSb.AppendLine(" <div class=\"meta\">שמירה על ידי: " + encUser + "</div>");
            projectSb.AppendLine(" </div>");

            if (!string.IsNullOrWhiteSpace(leaderDisplay))
            {
                projectSb.AppendLine(" <div class=\"leader\">• מתכנן מוביל: " + WebUtility.HtmlEncode(leaderDisplay) + "</div>");
            }

            if (!string.IsNullOrWhiteSpace(encNotes))
            {
                projectSb.AppendLine(" <div class=\"note\">" + encNotes + "</div>");
            }

            // Display full folder where everything was saved and link to the PDF
            projectSb.AppendLine(" <div class=\"muted\">תאריך שמירה: " + WebUtility.HtmlEncode(DateTime.Now.ToString("F")) + "</div>");
            projectSb.AppendLine(" <div style=\"margin-top:6px;font-size:12px;\">ההודעה נשמרה ב: " + WebUtility.HtmlEncode(saveFolder ?? string.Empty) + " &nbsp; <a class=\"file\" href=\"" + WebUtility.HtmlEncode(pdfHref) + "\" target=\"_blank\">(" + WebUtility.HtmlEncode(pdfFileName) + ")</a></div>");
            projectSb.AppendLine("</section>");

            writer.Write(projectSb.ToString());

            // Employees chunk
            var empSb = new StringBuilder();
            empSb.AppendLine("<section class=\"card\"> ");
            empSb.AppendLine(" <h2 style=\"font-size:14px;margin:8px0\">עובדים בפרויקט</h2>");
            empSb.AppendLine(" <table>");
            empSb.AppendLine(" <thead><tr><th>שם מלא</th><th>אימייל</th></tr></thead>");
            empSb.AppendLine(" <tbody>");

            if (employees != null && employees.Count >0)
            {
                foreach (var emp in employees)
                {
                    var name = WebUtility.HtmlEncode(((emp?.FirstName ?? string.Empty) + " " + (emp?.LastName ?? string.Empty)).Trim());
                    var email = WebUtility.HtmlEncode(emp?.EmailAddress ?? string.Empty);
                    empSb.AppendLine(" <tr><td>" + name + "</td><td>" + email + "</td></tr>");
                }
            }
            else
            {
                empSb.AppendLine(" <tr><td colspan=\"2\" style=\"text-align:center;color:#666;\">לא נמצאו עובדים</td></tr>");
            }

            empSb.AppendLine(" </tbody>");
            empSb.AppendLine(" </table>");
            empSb.AppendLine("</section>");

            writer.Write(empSb.ToString());

            // Attachments chunk
            var attSb = new StringBuilder();
            attSb.AppendLine("<section class=\"card\"> ");
            attSb.AppendLine(" <h2 style=\"font-size:14px;margin:8px0\">קבצים מצורפים</h2>");
            attSb.AppendLine(" <table>");
            attSb.AppendLine(" <thead><tr><th>קובץ</th><th>גודל</th></tr></thead>");
            attSb.AppendLine(" <tbody>");

            if (attachmentDisplayNames != null && attachmentDisplayNames.Count >0)
            {
                for (int i =0; i < attachmentDisplayNames.Count; i++)
                {
                    attSb.AppendLine(" <tr><td><a class=\"file\" href=\"" + attachmentHrefs[i] + "\" target=\"_blank\">" + attachmentDisplayNames[i] + "</a></td><td>" + attachmentSizes[i] + "</td></tr>");
                }
            }
            else
            {
                attSb.AppendLine(" <tr><td colspan=\"2\" style=\"text-align:center;color:#666;\">לא נבחרו קבצים לשמירה</td></tr>");
            }

            attSb.AppendLine(" </tbody>");
            attSb.AppendLine(" </table>");
            attSb.AppendLine("</section>");

            writer.Write(attSb.ToString());

            // Footer
            var footerSb = new StringBuilder();
            footerSb.AppendLine("</div>");
            footerSb.AppendLine("</body>");
            footerSb.AppendLine("</html>");

            writer.Write(footerSb.ToString());
        }
    }
}
