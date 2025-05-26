using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Drawing;
using System.Threading;
using System.Windows.Forms;

namespace SaveAsPDF.Tests
{
    [TestClass]
    public class XMessageBoxTests
    {
        private DialogResult result;
        private string loggedCaption;
        private string loggedMessage;

        [TestMethod]
        public void Show_DefaultBehavior()
        {
            RunOnUIThread(() =>
            {
                result = XMessageBox.Show("Default test message.");
            });

            Assert.AreEqual(DialogResult.OK, result);
        }

        [TestMethod]
        public void Show_OKCancel_Button()
        {
            RunOnUIThread(() =>
            {
                result = XMessageBox.Show(
                    "Test OKCancel",
                    "Test",
                    XMessageBoxButtons.OKCancel,
                    XMessageBoxIcon.Warning,
                    autoCloseMilliseconds: 15000);
            });

            Assert.AreEqual(DialogResult.Cancel, result); // Auto-close defaults to Cancel
        }

        [TestMethod]
        public void Show_YesNo_Button()
        {
            RunOnUIThread(() =>
            {
                result = XMessageBox.Show(
                    "Test YesNo",
                    "Test",
                    XMessageBoxButtons.YesNo,
                    XMessageBoxIcon.Question,
                    autoCloseMilliseconds: 15000);
            });

            Assert.AreEqual(DialogResult.No, result); // Auto-close defaults to No
        }

        [TestMethod]
        public void Show_CustomIcon()
        {
            RunOnUIThread(() =>
            {
                result = XMessageBox.Show(
                    "Custom icon test.",
                    "Icon Test",
                    XMessageBoxButtons.OK,
                    customIcon: SystemIcons.Shield.ToBitmap(),
                    autoCloseMilliseconds: 15000);
            });

            Assert.AreEqual(DialogResult.OK, result);
        }

        [TestMethod]
        public void Show_CenterAlignment()
        {
            RunOnUIThread(() =>
            {
                result = XMessageBox.Show(
                    "Center aligned message.",
                    "Alignment Test",
                    XMessageBoxButtons.OK,
                    XMessageBoxIcon.Information,
                    XMessageAlignment.Center,
                    autoCloseMilliseconds: 15000);
            });

            Assert.AreEqual(DialogResult.OK, result);
        }

        [TestMethod]
        public void Show_HebrewLanguage_OK()
        {
            RunOnUIThread(() =>
            {
                result = XMessageBox.Show(
                    "בדיקה בעברית",
                    "כותרת",
                    XMessageBoxButtons.OK,
                    XMessageBoxIcon.Information,
                    XMessageAlignment.Right,
                    XMessageLanguage.Hebrew,
                    autoCloseMilliseconds: 15000);
            });

            Assert.AreEqual(DialogResult.OK, result);
        }

        [TestMethod]
        public void Show_HebrewLanguage_OKCancel()
        {
            RunOnUIThread(() =>
            {
                result = XMessageBox.Show(
                    "בדיקת אישור וביטול בעברית",
                    "כותרת",
                    XMessageBoxButtons.OKCancel,
                    XMessageBoxIcon.Warning,
                    XMessageAlignment.Right,
                    XMessageLanguage.Hebrew,
                    autoCloseMilliseconds: 15000);
            });

            Assert.AreEqual(DialogResult.Cancel, result);
        }

        [TestMethod]
        public void Show_HebrewLanguage_YesNo()
        {
            RunOnUIThread(() =>
            {
                result = XMessageBox.Show(
                    "בדיקת כן ולא בעברית",
                    "כותרת",
                    XMessageBoxButtons.YesNo,
                    XMessageBoxIcon.Question,
                    XMessageAlignment.Right,
                    XMessageLanguage.Hebrew,
                    autoCloseMilliseconds: 15000);
            });

            Assert.AreEqual(DialogResult.No, result);
        }

        [TestMethod]
        public void Show_HebrewLanguage_YesNoCancel()
        {
            RunOnUIThread(() =>
            {
                result = XMessageBox.Show(
                    "בדיקת כן, לא וביטול בעברית",
                    "כותרת",
                    XMessageBoxButtons.YesNoCancel,
                    XMessageBoxIcon.Question,
                    XMessageAlignment.Right,
                    XMessageLanguage.Hebrew,
                    autoCloseMilliseconds: 15000);
            });

            Assert.AreEqual(DialogResult.Cancel, result);
        }

        [TestMethod]
        public void Show_HebrewLanguage_RetryCancel()
        {
            RunOnUIThread(() =>
            {
                result = XMessageBox.Show(
                    "בדיקת נסה שוב וביטול בעברית",
                    "כותרת",
                    XMessageBoxButtons.RetryCancel,
                    XMessageBoxIcon.Error,
                    XMessageAlignment.Right,
                    XMessageLanguage.Hebrew,
                    autoCloseMilliseconds: 15000);
            });

            Assert.AreEqual(DialogResult.Cancel, result);
        }

        [TestMethod]
        public void Show_CustomFontAndBackground()
        {
            RunOnUIThread(() =>
            {
                result = XMessageBox.Show(
                    "Custom font and background test.",
                    "Custom Test",
                    XMessageBoxButtons.OK,
                    XMessageBoxIcon.Warning,
                    customFont: new Font("Arial", 12, FontStyle.Italic),
                    backColorOverride: Color.LightBlue,
                    autoCloseMilliseconds: 15000);
            });

            Assert.AreEqual(DialogResult.OK, result);
        }

        [TestMethod]
        public void Show_AutoClose()
        {
            RunOnUIThread(() =>
            {
                result = XMessageBox.Show(
                    "Auto-close test.",
                    "AutoClose",
                    XMessageBoxButtons.OK,
                    XMessageBoxIcon.Information,
                    autoCloseMilliseconds: 15000);
            });

            Assert.AreEqual(DialogResult.OK, result);
        }

        [TestMethod]
        public void Show_WithLogging()
        {
            RunOnUIThread(() =>
            {
                result = XMessageBox.Show(
                    "Logging test message.",
                    "Log Caption",
                    XMessageBoxButtons.OK,
                    XMessageBoxIcon.Information,
                    logCallback: (caption, message) =>
                    {
                        loggedCaption = caption;
                        loggedMessage = message;
                    },
                    autoCloseMilliseconds: 15000);
            });

            Assert.AreEqual("Log Caption", loggedCaption);
            Assert.AreEqual("Logging test message.", loggedMessage);
        }

        [TestMethod]
        public void Show_NoCentering()
        {
            RunOnUIThread(() =>
            {
                result = XMessageBox.Show(
                    "Manual position test.",
                    "Position",
                    XMessageBoxButtons.OK,
                    XMessageBoxIcon.Information,
                    centerToActiveWindow: false,
                    autoCloseMilliseconds: 15000);
            });

            Assert.AreEqual(DialogResult.OK, result);
        }

        /// <summary>
        /// Ensures all UI code runs on a WinForms-compatible thread.
        /// </summary>
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

            thread.SetApartmentState(ApartmentState.STA); // Required for WinForms
            thread.Start();
            thread.Join();

            if (caught != null)
                throw caught;
        }
    }
}
