using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Windows.Forms;

namespace SaveAsPDF.Tests
{
    [TestClass]
    public class XMessageBoxTests
    {
        [TestMethod]
        public void Show_1ShouldReturnOK_WhenOKButtonIsClicked()
        {
            // Arrange
            //string text = "שורה בעברית טיפה ארוכה";
            string text = "Hello! I'm Jithu Thomas, a passionate .NET Developer with a knack for creating robust and scalable";
            string caption = "Test caption";
            XMessageBoxButtons buttons = XMessageBoxButtons.OKCancel;
            XMessageBoxIcon icon = XMessageBoxIcon.Information;
            XMessageAlignment alignment = XMessageAlignment.Left;
            XMessageLanguage language = XMessageLanguage.English;


            // Act
            DialogResult result = XMessageBox.Show(text, caption, buttons, icon, alignment, language);
            //DialogResult result = XMessageBox.Show(text, caption, buttons, icon);
            // Assert
            Assert.AreEqual(DialogResult.OK, result);
        }

        [TestMethod]
        public void Show_2ShouldReturnCancel_WhenCancelButtonIsClicked()
        {
            // Arrange
            string text = "hit cancel";
            string caption = "Test caption";
            XMessageBoxButtons buttons = XMessageBoxButtons.OKCancel;
            XMessageBoxIcon icon = XMessageBoxIcon.Warning;
            XMessageAlignment alignment = XMessageAlignment.Left;
            XMessageLanguage language = XMessageLanguage.English;
            // Act
            DialogResult result = XMessageBox.Show(text, caption, buttons, icon, alignment, language);

            // Assert
            Assert.AreEqual(DialogResult.Cancel, result);
        }

        [TestMethod]
        public void Show_3ShouldReturnYes_WhenYesButtonIsClicked()
        {
            // Arrange
            string text = "hit yes";
            string caption = "Test caption";
            XMessageBoxButtons buttons = XMessageBoxButtons.YesNo;
            XMessageBoxIcon icon = XMessageBoxIcon.Question;

            // Act
            DialogResult result = XMessageBox.Show(text, caption, buttons, icon);

            // Assert
            Assert.AreEqual(DialogResult.Yes, result);
        }

        [TestMethod]
        public void Show_4ShouldReturnNo_WhenNoButtonIsClicked()
        {
            // Arrange
            string text = "hit no";
            string caption = "Test caption";
            XMessageBoxButtons buttons = XMessageBoxButtons.YesNo;
            XMessageBoxIcon icon = XMessageBoxIcon.Error;

            // Act
            DialogResult result = XMessageBox.Show(text, caption, buttons, icon);

            // Assert
            Assert.AreEqual(DialogResult.No, result);
        }
    }
}
