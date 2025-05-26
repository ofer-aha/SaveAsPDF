using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Threading;
using System.Windows.Forms;

namespace SaveAsPDF.Tests
{
    [TestClass]
    public class XMessageBoxInputTests
    {
        private string loggedCaption;
        private string loggedMessage;

        [TestMethod]
        public void ShowInput_DefaultBehavior()
        {
            SaveAsPDF.InputBoxResult result = null;
            RunOnUIThread(() =>
            {
                result = SaveAsPDF.XMessageBox.ShowInput("Enter value:", "Input Test", "default");
            });
            Assert.IsNotNull(result);
            Assert.AreEqual("default", result.Text);
        }

        [TestMethod]
        public void ShowInput_HebrewLanguage()
        {
            SaveAsPDF.InputBoxResult result = null;
            RunOnUIThread(() =>
            {
                result = SaveAsPDF.XMessageBox.ShowInput("הכנס ערך:", "קלט", "ברירת מחדל", null, -1, -1, SaveAsPDF.XMessageLanguage.Hebrew);
            });
            Assert.IsNotNull(result);
            Assert.AreEqual("ברירת מחדל", result.Text);
        }

        [TestMethod]
        public void ShowInput_ValidationBlocksInvalidInput()
        {
            bool validatorCalled = false;
            SaveAsPDF.InputBoxResult result = null;
            SaveAsPDF.InputBoxValidatingHandler validator = (sender, e) =>
            {
                validatorCalled = true;
                if (string.IsNullOrWhiteSpace(e.Text))
                {
                    e.Cancel = true;
                    e.Message = "Input required";
                }
            };
            RunOnUIThread(() =>
            {
                result = SaveAsPDF.XMessageBox.ShowInput("Prompt", "Title", "", validator);
            });
            Assert.IsNotNull(result);
            Assert.IsTrue(validatorCalled);
        }

        [TestMethod]
        public void ShowInput_Positioned()
        {
            SaveAsPDF.InputBoxResult result = null;
            RunOnUIThread(() =>
            {
                result = SaveAsPDF.XMessageBox.ShowInput("Prompt", "Title", "test", null, 100, 200);
            });
            Assert.IsNotNull(result);
            Assert.AreEqual("test", result.Text);
        }

        private void RunOnUIThread(Action action)
        {
            Exception caught = null;
            var thread = new Thread(() =>
            {
                try
                {
                    Application.EnableVisualStyles();
                    action();
                }
                catch (Exception ex)
                {
                    caught = ex;
                }
            });
            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();
            thread.Join();
            if (caught != null)
                throw caught;
        }
    }
}