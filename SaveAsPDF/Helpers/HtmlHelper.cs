using System;
using System.Collections.Generic;
using System.IO;

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
            string userName)
        {
            // Sanitize the project name using SafeFolderName
            string sanitizedProjectName = projectName.SafeFolderName();

            // Build the HTML table style and header
            string tableStyle = @"
            <html>
            <head>
                <style>
                    table, td, th {
                        border-collapse: collapse;
                        border: 1px solid black;
                        padding: 4px;
                        margin: 0;
                    }
                    th {
                        background-color: #f2f2f2;
                        text-align: center;
                        font-size: 12px;
                    }
                    td {
                        text-align: right;
                        font-size: 12px;
                    }
                    tr:nth-child(even) {
                        background-color: #f9f9f9;
                    }
                    a {
                        color: #007bff;
                        text-decoration: underline;
                    }
                </style>
            </head>
            <body>";

            // Project metadata table
            string pdfFileName = $"{sanitizedProjectName}.pdf";
            string pdfFullPath = Path.Combine(sPath, pdfFileName);
            string projectData = $@"
            <table style='width: 100%; direction: rtl; margin-bottom: 10px;'>
                <thead>
                    <tr>
                        <th>שם הפרויקט</th>
                        <th>מס' פרויקט</th>
                        <th>הערות</th>
                    </tr>
                </thead>
                <tbody>
                    <tr>
                        <td>{projectName}</td>
                        <td>{projectID}</td>
                        <td>{notes.Replace(Environment.NewLine, "<br>")}</td>
                    </tr>
                    <tr>
                        <td colspan='3'>שם משתמש: {userName}</td>
                    </tr>
                    <tr>
                        <td colspan='3'>תאריך השמירה: {DateTime.Now:F}</td>
                    </tr>
                    <tr>
                        <td colspan='3'>
                            ההודעה נשמרה ב: 
                            <a href='file:///{pdfFullPath.Replace("\\", "/")}' target='_blank'>{pdfFileName}</a>
                        </td>
                    </tr>
                </tbody>
            </table>";

            // Employee table
            string employeeTable = @"
            <table style='width: 100%; direction: rtl; margin-bottom: 10px;'>
                <thead>
                    <tr>
                        <th>שם מלא</th>
                        <th>אימייל</th>
                    </tr>
                </thead>
                <tbody>";

            foreach (var employee in employees)
            {
                employeeTable += $@"
                    <tr>
                        <td>{employee.FirstName} {employee.LastName}</td>
                        <td>{employee.EmailAddress}</td>
                    </tr>";
            }

            employeeTable += @"
                </tbody>
            </table>";

            // Attachment table with links
            string attachmentTable = @"
            <table style='width: 100%; direction: rtl; margin-bottom: 10px;'>
                <thead>
                    <tr>
                        <th>שם קובץ (נתיב מלא)</th>
                        <th>גודל קובץ</th>
                    </tr>
                </thead>
                <tbody>";

            foreach (var attachment in attachments)
            {
                string sanitizedFileName = attachment.fileName.SafeFolderName();
                string fullPath = Path.Combine(sPath, sanitizedFileName);
                string fileLink = $"<a href='file:///{fullPath.Replace("\\", "/")}' target='_blank'>{sanitizedFileName}</a>";
                attachmentTable += $@"
                    <tr>
                        <td>{fileLink}</td>
                        <td>{attachment.fileSize}</td>
                    </tr>";
            }

            attachmentTable += @"
                </tbody>
            </table>";

            // Combine all sections and close HTML structure
            return tableStyle + projectData + employeeTable + attachmentTable + "</body></html>";
        }
    }
}
