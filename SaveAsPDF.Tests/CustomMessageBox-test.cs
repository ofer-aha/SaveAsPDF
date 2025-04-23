using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Windows.Forms;

namespace SaveAsPDF.Tests
{
    [TestClass]
    public class CustomMessageBoxTests
    {
        [TestMethod]
        public void Show_1ShouldReturnOK_WhenOKButtonIsClicked()
        {
            // Arrange
            //string text = "שורה בעברית טיפה ארוכה";
            string text = "Hello! I'm Jithu Thomas, a passionate .NET Developer with a knack for creating robust and scalable";
            string caption = "Test caption";
            CustomMessageBoxButtons buttons = CustomMessageBoxButtons.OKCancel;
            CustomMessageBoxIcon icon = CustomMessageBoxIcon.Information;
            CustomMessageAlignment alignment = CustomMessageAlignment.Left;
            CustomMessageLanguage language = CustomMessageLanguage.English;


            // Act
            DialogResult result = CustomMessageBox.Show(text, caption, buttons, icon, alignment, language);

            // Assert
            Assert.AreEqual(DialogResult.OK, result);
        }

        [TestMethod]
        public void Show_2ShouldReturnCancel_WhenCancelButtonIsClicked()
        {
            // Arrange
            string text = "hit cancel";
            string caption = "Test caption";
            CustomMessageBoxButtons buttons = CustomMessageBoxButtons.OKCancel;
            CustomMessageBoxIcon icon = CustomMessageBoxIcon.Warning;
            CustomMessageAlignment alignment = CustomMessageAlignment.Left;
            CustomMessageLanguage language = CustomMessageLanguage.English;
            // Act
            DialogResult result = CustomMessageBox.Show(text, caption, buttons, icon, alignment, language);

            // Assert
            Assert.AreEqual(DialogResult.Cancel, result);
        }

        [TestMethod]
        public void Show_3ShouldReturnYes_WhenYesButtonIsClicked()
        {
            // Arrange
            string text = "hit yes";
            string caption = "Test caption";
            CustomMessageBoxButtons buttons = CustomMessageBoxButtons.YesNo;
            CustomMessageBoxIcon icon = CustomMessageBoxIcon.Question;

            // Act
            DialogResult result = CustomMessageBox.Show(text, caption, buttons, icon);

            // Assert
            Assert.AreEqual(DialogResult.Yes, result);
        }

        [TestMethod]
        public void Show_4ShouldReturnNo_WhenNoButtonIsClicked()
        {
            // Arrange
            string text = "hit no";
            string caption = "Test caption";
            CustomMessageBoxButtons buttons = CustomMessageBoxButtons.YesNo;
            CustomMessageBoxIcon icon = CustomMessageBoxIcon.Error;

            // Act
            DialogResult result = CustomMessageBox.Show(text, caption, buttons, icon);

            // Assert
            Assert.AreEqual(DialogResult.No, result);
        }
    }
}
