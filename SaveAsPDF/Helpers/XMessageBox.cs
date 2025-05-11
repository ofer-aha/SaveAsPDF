using System;
using System.Drawing;
using System.Windows.Forms;

namespace SaveAsPDF
{
    /// <summary>
    /// Provides a static method to display a custom message box with various configurations.
    /// </summary>
    public static class XMessageBox
    {
        /// <summary>
        /// Displays a custom message box with the specified parameters.
        /// </summary>
        /// <param name="text">The message text to display in the message box.</param>
        /// <param name="caption">The caption text for the message box title bar. Default is an empty string.</param>
        /// <param name="buttons">The buttons to display in the message box (e.g., OK, OKCancel, YesNo). Default is <see cref="XMessageBoxButtons.OK"/>.</param>
        /// <param name="icon">The icon to display in the message box (e.g., Information, Warning, Error). Default is <see cref="XMessageBoxIcon.Information"/>.</param>
        /// <param name="alignment">The alignment of the message text (e.g., Left, Center, Right). Default is <see cref="XMessageAlignment.Left"/>.</param>
        /// <param name="language">The language of the message box buttons (e.g., English, Hebrew). Default is <see cref="XMessageLanguage.English"/>.</param>
        /// <param name="useTextBox">Indicates whether to display the message in a multi-line text box. Default is <c>false</c>.</param>
        /// <param name="isDarkTheme">Indicates whether to use a dark theme for the message box. Default is <c>false</c>.</param>
        /// <param name="customFont">A custom font to use for the message text. Default is <c>null</c>, which uses the default font.</param>
        /// <param name="backColorOverride">An optional background color override for the message box. Default is <c>null</c>.</param>
        /// <param name="autoCloseMilliseconds">The time in milliseconds after which the message box will automatically close. Default is <c>0</c> (no auto-close).</param>
        /// <param name="customIcon">An optional custom icon to display in the message box. Default is <c>null</c>.</param>
        /// <param name="logCallback">An optional callback to log the caption and message text. Default is <c>null</c>.</param>
        /// <param name="centerToActiveWindow">Indicates whether to center the message box to the active window. Default is <c>true</c>.</param>
        /// <returns>The <see cref="DialogResult"/> indicating which button was clicked by the user.</returns>
        public static DialogResult Show(
            string text,
            string caption = "",
            XMessageBoxButtons buttons = XMessageBoxButtons.OK,
            XMessageBoxIcon icon = XMessageBoxIcon.Information,
            XMessageAlignment alignment = XMessageAlignment.Left,
            XMessageLanguage language = XMessageLanguage.English,
            bool useTextBox = false,
            bool isDarkTheme = false,
            Font customFont = null,
            Color? backColorOverride = null,
            int autoCloseMilliseconds = 0,
            Image customIcon = null,
            Action<string, string> logCallback = null,
            bool centerToActiveWindow = true)
        {
            // Optional logging
            logCallback?.Invoke(caption, text);

            using (var form = new XMessageBoxForm(
                text,
                caption,
                buttons,
                icon,
                alignment,
                language,
                useTextBox,
                isDarkTheme,
                customFont,
                backColorOverride,
                autoCloseMilliseconds,
                customIcon,
                centerToActiveWindow))
            {
                return form.ShowDialog();
            }
        }
    }
}
