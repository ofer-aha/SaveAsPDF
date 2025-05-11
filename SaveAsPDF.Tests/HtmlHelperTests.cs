using Microsoft.VisualStudio.TestTools.UnitTesting;
using SaveAsPDF.Helpers;
using SaveAsPDF.Models;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;

namespace SaveAsPDF.Tests
{
    [TestClass]
    public class HtmlHelperTests
    {
        private void OpenHtmlInBrowser(string htmlContent)
        {
            // Save the HTML content to a temporary file
            string tempFilePath = Path.Combine(Path.GetTempPath(), "TestOutput.html");
            File.WriteAllText(tempFilePath, htmlContent);

            // Open the file in the default browser
            Process.Start(new ProcessStartInfo
            {
                FileName = tempFilePath,
                UseShellExecute = true
            });
        }

        [TestMethod]
        public void GenerateHtmlContent_ValidInput_ReturnsHtmlString()
        {
            // Arrange
            string sPath = @"C:\TestPath";
            var employees = new List<EmployeeModel>
            {
                new EmployeeModel { FirstName = "John", LastName = "Doe", EmailAddress = "john.doe@example.com" },
                new EmployeeModel { FirstName = "Jane", LastName = "Smith", EmailAddress = "jane.smith@example.com" }
            };
            var attachments = new List<AttachmentsModel>
            {
                new AttachmentsModel { fileName = "file1.txt", fileSize = "10 KB" },
                new AttachmentsModel { fileName = "file2.pdf", fileSize = "20 KB" }
            };
            string projectName = "Test Project";
            string projectID = "12345";
            string notes = "This is a test project.";
            string userName = "TestUser";

            // Act
            string result = HtmlHelper.GenerateHtmlContent(sPath, employees, attachments, projectName, projectID, notes, userName);

            // Assert
            Assert.IsNotNull(result);
            Assert.IsTrue(result.Contains("Test Project"));
            Assert.IsTrue(result.Contains("12345"));
            Assert.IsTrue(result.Contains("This is a test project."));
            Assert.IsTrue(result.Contains("TestUser"));
            Assert.IsTrue(result.Contains("John Doe"));
            Assert.IsTrue(result.Contains("jane.smith@example.com"));
            Assert.IsTrue(result.Contains("file1.txt"));
            Assert.IsTrue(result.Contains("10 KB"));

            // Validate the new rows
            Assert.IsTrue(result.Contains("תאריך השמירה"));
            Assert.IsTrue(result.Contains("ההודעה נשמרה ב"));
            Assert.IsTrue(result.Contains(Path.Combine(sPath, $"{projectName}.html")));

            // Open the result in the default browser
            OpenHtmlInBrowser(result);
        }

        [TestMethod]
        public void GenerateHtmlContent_EmptyEmployeesAndAttachments_ReturnsHtmlString()
        {
            // Arrange
            string sPath = @"C:\TestPath";
            var employees = new List<EmployeeModel>();
            var attachments = new List<AttachmentsModel>();
            string projectName = "Empty Test Project";
            string projectID = "00000";
            string notes = "No employees or attachments.";
            string userName = "EmptyUser";

            // Act
            string result = HtmlHelper.GenerateHtmlContent(sPath, employees, attachments, projectName, projectID, notes, userName);

            // Assert
            Assert.IsNotNull(result);
            Assert.IsTrue(result.Contains("Empty Test Project"));
            Assert.IsTrue(result.Contains("00000"));
            Assert.IsTrue(result.Contains("No employees or attachments."));
            Assert.IsTrue(result.Contains("EmptyUser"));

            // Validate the new rows
            Assert.IsTrue(result.Contains("תאריך השמירה"));
            Assert.IsTrue(result.Contains("ההודעה נשמרה ב"));
            Assert.IsTrue(result.Contains(Path.Combine(sPath, $"{projectName}.html")));

            // Open the result in the default browser
            OpenHtmlInBrowser(result);
        }

        [TestMethod]
        public void GenerateHtmlContent_SpecialCharactersInInput_ReturnsEscapedHtmlString()
        {
            // Arrange
            string sPath = @"C:\TestPath";
            var employees = new List<EmployeeModel>
            {
                new EmployeeModel { FirstName = "John & Jane", LastName = "Doe <Smith>", EmailAddress = "john.doe@example.com" }
            };
            var attachments = new List<AttachmentsModel>
            {
                new AttachmentsModel { fileName = "file1&2.txt", fileSize = "10 KB" }
            };
            string projectName = "Test <Project>";
            string projectID = "123&45";
            string notes = "This is a test project with special characters: <, >, &.";
            string userName = "TestUser";

            // Act
            string result = HtmlHelper.GenerateHtmlContent(sPath, employees, attachments, projectName, projectID, notes, userName);

            // Assert
            Assert.IsNotNull(result);
            Assert.IsTrue(result.Contains("Test <Project>"));
            Assert.IsTrue(result.Contains("123&45"));
            Assert.IsTrue(result.Contains("This is a test project with special characters: <, >, &."));
            Assert.IsTrue(result.Contains("John & Jane"));
            Assert.IsTrue(result.Contains("Doe <Smith>"));
            Assert.IsTrue(result.Contains("file1&2.txt"));

            // Validate the new rows
            Assert.IsTrue(result.Contains("תאריך השמירה"));
            Assert.IsTrue(result.Contains("ההודעה נשמרה ב"));
            Assert.IsTrue(result.Contains(Path.Combine(sPath, $"{projectName}.html")));

            // Open the result in the default browser
            OpenHtmlInBrowser(result);
        }

        [TestMethod]
        public void GenerateHtmlContent_LongInput_ReturnsHtmlString()
        {
            // Arrange
            string sPath = @"C:\TestPath";
            var employees = new List<EmployeeModel>();
            for (int i = 0; i < 100; i++)
            employees.Add(new EmployeeModel { FirstName = $"FirstName{i}", LastName = $"LastName{i}", EmailAddress = $"email{i}@example.com" });

            var attachments = new List<AttachmentsModel>();
            for (int i = 0; i < 50; i++)
            attachments.Add(new AttachmentsModel { fileName = $"file{i}.txt", fileSize = $"{i * 10} KB" });

            string projectName = "Long Test Project";
            string projectID = "99999";
            string notes = "This is a test project with a lot of data.";
            string userName = "LongUser";

            // Act
            string result = HtmlHelper.GenerateHtmlContent(sPath, employees, attachments, projectName, projectID, notes, userName);

            // Assert
            Assert.IsNotNull(result);
            Assert.IsTrue(result.Contains("Long Test Project"));
            Assert.IsTrue(result.Contains("99999"));
            Assert.IsTrue(result.Contains("This is a test project with a lot of data."));
            Assert.IsTrue(result.Contains("LongUser"));
            Assert.IsTrue(result.Contains("FirstName0"));
            Assert.IsTrue(result.Contains("LastName99"));
            Assert.IsTrue(result.Contains("file49.txt"));
            Assert.IsTrue(result.Contains("490 KB"));

            // Validate the new rows
            Assert.IsTrue(result.Contains("תאריך השמירה"));
            Assert.IsTrue(result.Contains("ההודעה נשמרה ב"));
            Assert.IsTrue(result.Contains(Path.Combine(sPath, $"{projectName}.html")));

            // Open the result in the default browser
            OpenHtmlInBrowser(result);
        }
    }
}
