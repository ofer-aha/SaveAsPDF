using System;
using System.Drawing;
using System.Windows.Forms;
using System.ComponentModel;

namespace SaveAsPDF
{
    /// <summary>
    /// Specifies the buttons to display in the message box.
    /// </summary>
    public enum XMessageBoxButtons
    {
        OK,
        OKCancel,
        YesNo,
        YesNoCancel,
        RetryCancel
    }

    /// <summary>
    /// Specifies the icon to display in the message box.
    /// </summary>
    public enum XMessageBoxIcon
    {
        None,
        Information,
        Warning,
        Error,
        Question
    }

    /// <summary>
    /// Specifies the alignment of the message text.
    /// </summary>
    public enum XMessageAlignment
    {
        Left,
        Center,
        Right
    }

    /// <summary>
    /// Specifies the language of the message box buttons.
    /// </summary>
    public enum XMessageLanguage
    {
        Hebrew,
        English
    }

    /// <summary>
    /// Result for input box dialog.
    /// </summary>
    public class InputBoxResult
    {
        public bool OK { get; set; }
        public string Text { get; set; }
    }

    /// <summary>
    /// EventArgs for input box validation.
    /// </summary>
    public class InputBoxValidatingArgs : EventArgs
    {
        public string Text { get; set; }
        public string Message { get; set; }
        public bool Cancel { get; set; }
    }

    /// <summary>
    /// Delegate for input box validation.
    /// </summary>
    public delegate void InputBoxValidatingHandler(object sender, InputBoxValidatingArgs e);

    /// <summary>
    /// Provides static methods to display a custom message box and input box with various configurations.
    /// </summary>
    public static partial class XMessageBox
    {
        /// <summary>
        /// Displays a custom message box with the specified parameters.
        /// For Hebrew, text is right-aligned, buttons are always center-aligned.
        /// If autoCloseMilliseconds > 0, the title displays a countdown.
        /// </summary>
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
            logCallback?.Invoke(caption, text);

            using (Form form = new Form())
            {
                bool isHebrew = language == XMessageLanguage.Hebrew;
                form.FormBorderStyle = FormBorderStyle.FixedDialog;
                form.MaximizeBox = false;
                form.MinimizeBox = false;
                form.StartPosition = centerToActiveWindow ? FormStartPosition.CenterParent : FormStartPosition.CenterScreen;
                form.ClientSize = new Size(400, useTextBox ? 200 : 140);
                form.ShowInTaskbar = false;
                form.RightToLeft = isHebrew ? RightToLeft.Yes : RightToLeft.No;
                form.RightToLeftLayout = isHebrew;

                if (isDarkTheme)
                {
                    form.BackColor = Color.FromArgb(40, 40, 40);
                    form.ForeColor = Color.White;
                }
                if (backColorOverride.HasValue)
                {
                    form.BackColor = backColorOverride.Value;
                }

                // Icon
                PictureBox iconBox = null;
                if (icon != XMessageBoxIcon.None || customIcon != null)
                {
                    iconBox = new PictureBox
                    {
                        Size = new Size(48, 48),
                        Location = new Point(15, 20),
                        SizeMode = PictureBoxSizeMode.StretchImage
                    };
                    if (customIcon != null)
                    {
                        iconBox.Image = customIcon;
                    }
                    else
                    {
                        switch (icon)
                        {
                            case XMessageBoxIcon.Information:
                                iconBox.Image = SystemIcons.Information.ToBitmap();
                                break;
                            case XMessageBoxIcon.Warning:
                                iconBox.Image = SystemIcons.Warning.ToBitmap();
                                break;
                            case XMessageBoxIcon.Error:
                                iconBox.Image = SystemIcons.Error.ToBitmap();
                                break;
                            case XMessageBoxIcon.Question:
                                iconBox.Image = SystemIcons.Question.ToBitmap();
                                break;
                            default:
                                iconBox.Image = null;
                                break;
                        }
                    }
                }

                // Message
                Control messageControl;
                int messageLeft = iconBox != null ? 75 : 15;
                int messageWidth = form.ClientSize.Width - (iconBox != null ? 90 : 30);

                if (useTextBox)
                {
                    var textBox = new TextBox
                    {
                        Multiline = true,
                        ReadOnly = true,
                        BorderStyle = BorderStyle.None,
                        Location = new Point(messageLeft, 20),
                        Size = new Size(messageWidth, 80),
                        Text = text,
                        BackColor = form.BackColor,
                        ForeColor = form.ForeColor,
                        TabStop = false,
                        TextAlign = isHebrew ? HorizontalAlignment.Right :
                            alignment == XMessageAlignment.Center ? HorizontalAlignment.Center :
                            alignment == XMessageAlignment.Right ? HorizontalAlignment.Right : HorizontalAlignment.Left,
                        RightToLeft = isHebrew ? RightToLeft.Yes : RightToLeft.No
                    };
                    if (customFont != null)
                        textBox.Font = customFont;
                    messageControl = textBox;
                }
                else
                {
                    var label = new Label
                    {
                        AutoSize = false,
                        Location = new Point(messageLeft, 20),
                        Size = new Size(messageWidth, 80),
                        Text = text,
                        BackColor = Color.Transparent,
                        ForeColor = form.ForeColor,
                        TextAlign = isHebrew ? ContentAlignment.MiddleRight :
                            alignment == XMessageAlignment.Center ? ContentAlignment.MiddleCenter :
                            alignment == XMessageAlignment.Right ? ContentAlignment.MiddleRight : ContentAlignment.MiddleLeft,
                        RightToLeft = isHebrew ? RightToLeft.Yes : (alignment == XMessageAlignment.Right ? RightToLeft.Yes : RightToLeft.No)
                    };
                    if (customFont != null)
                        label.Font = customFont;
                    messageControl = label;
                }

                // Buttons
                string[] buttonTexts;
                DialogResult[] buttonResults;
                int defaultButtonIndex = 0;

                switch (buttons)
                {
                    case XMessageBoxButtons.OKCancel:
                        buttonTexts = isHebrew ? new[] { "ביטול", "אישור" } : new[] { "OK", "Cancel" };
                        buttonResults = isHebrew ? new[] { DialogResult.Cancel, DialogResult.OK } : new[] { DialogResult.OK, DialogResult.Cancel };
                        defaultButtonIndex = isHebrew ? 1 : 0;
                        break;
                    case XMessageBoxButtons.YesNo:
                        buttonTexts = isHebrew ? new[] { "לא", "כן" } : new[] { "Yes", "No" };
                        buttonResults = isHebrew ? new[] { DialogResult.No, DialogResult.Yes } : new[] { DialogResult.Yes, DialogResult.No };
                        defaultButtonIndex = isHebrew ? 1 : 0;
                        break;
                    case XMessageBoxButtons.YesNoCancel:
                        buttonTexts = isHebrew ? new[] { "ביטול", "לא", "כן" } : new[] { "Yes", "No", "Cancel" };
                        buttonResults = isHebrew ? new[] { DialogResult.Cancel, DialogResult.No, DialogResult.Yes } : new[] { DialogResult.Yes, DialogResult.No, DialogResult.Cancel };
                        defaultButtonIndex = isHebrew ? 2 : 0;
                        break;
                    case XMessageBoxButtons.RetryCancel:
                        buttonTexts = isHebrew ? new[] { "ביטול", "נסה שוב" } : new[] { "Retry", "Cancel" };
                        buttonResults = isHebrew ? new[] { DialogResult.Cancel, DialogResult.Retry } : new[] { DialogResult.Retry, DialogResult.Cancel };
                        defaultButtonIndex = isHebrew ? 1 : 0;
                        break;
                    default:
                        buttonTexts = isHebrew ? new[] { "אישור" } : new[] { "OK" };
                        buttonResults = new[] { DialogResult.OK };
                        defaultButtonIndex = 0;
                        break;
                }

                int buttonWidth = 90;
                int buttonHeight = 30;
                int buttonSpacing = 15;
                int totalButtonsWidth = buttonTexts.Length * buttonWidth + (buttonTexts.Length - 1) * buttonSpacing;
                int buttonY = form.ClientSize.Height - buttonHeight - 20;
                int buttonX = (form.ClientSize.Width - totalButtonsWidth) / 2;

                var buttonsArr = new Button[buttonTexts.Length];
                for (int i = 0; i < buttonTexts.Length; i++)
                {
                    var btn = new Button
                    {
                        Text = buttonTexts[i],
                        Size = new Size(buttonWidth, buttonHeight),
                        Location = new Point(buttonX + i * (buttonWidth + buttonSpacing), buttonY),
                        DialogResult = buttonResults[i],
                        TextAlign = ContentAlignment.MiddleCenter,
                        RightToLeft = isHebrew ? RightToLeft.Yes : RightToLeft.No
                    };
                    if (isDarkTheme)
                    {
                        btn.BackColor = Color.FromArgb(60, 60, 60);
                        btn.ForeColor = Color.White;
                    }
                    form.Controls.Add(btn);
                    buttonsArr[i] = btn;
                }
                form.AcceptButton = buttonsArr[defaultButtonIndex];
                form.CancelButton = isHebrew ? buttonsArr[0] : buttonsArr[buttonsArr.Length - 1];

                if (iconBox != null)
                    form.Controls.Add(iconBox);
                form.Controls.Add(messageControl);

                // Auto-close timer: update title with countdown
                Timer timer = null;
                int remainingSeconds = 0;
                string baseTitle = caption ?? string.Empty;
                if (autoCloseMilliseconds > 0)
                {
                    remainingSeconds = (int)Math.Ceiling(autoCloseMilliseconds / 1000.0);
                    UpdateTitleWithTimer(form, baseTitle, language, remainingSeconds);
                    timer = new Timer { Interval = 1000 };
                    timer.Tick += (s, e) =>
                    {
                        remainingSeconds--;
                        UpdateTitleWithTimer(form, baseTitle, language, remainingSeconds);
                        if (remainingSeconds <= 0)
                        {
                            timer.Stop();
                            buttonsArr[defaultButtonIndex].PerformClick();
                        }
                    };
                    timer.Start();
                }
                else
                {
                    form.Text = baseTitle;
                }

                var result = form.ShowDialog();
                timer?.Dispose();
                return result;
            }
        }

        /// <summary>
        /// Updates the form title to include the auto-close timer in the appropriate language.
        /// </summary>
        private static void UpdateTitleWithTimer(Form form, string baseTitle, XMessageLanguage language, int seconds)
        {
            string timerMsg;
            if (language == XMessageLanguage.Hebrew)
                timerMsg = $" (סוגר בעוד: {seconds} {(seconds == 1 ? "שנייה" : "שניות")})";
            else
                timerMsg = $" (Closing in: {seconds} second{(seconds == 1 ? "" : "s")})";
            form.Text = baseTitle + timerMsg;
        }

        /// <summary>
        /// Displays an input box dialog with validation and returns the result.
        /// The input TextBox is right-aligned in Hebrew.
        /// </summary>
        public static InputBoxResult ShowInput(
            string prompt,
            string title,
            string defaultResponse = "",
            InputBoxValidatingHandler validator = null,
            int xPos = -1,
            int yPos = -1,
            XMessageLanguage language = XMessageLanguage.English)
        {
            using (Form form = new Form())
            {
                form.Text = title;
                form.FormBorderStyle = FormBorderStyle.FixedDialog;
                form.MaximizeBox = false;
                form.MinimizeBox = false;
                form.ClientSize = new Size(464, 104);
                form.StartPosition = (xPos >= 0 && yPos >= 0) ? FormStartPosition.Manual : FormStartPosition.CenterScreen;
                if (xPos >= 0 && yPos >= 0)
                {
                    form.Left = xPos;
                    form.Top = yPos;
                }
                bool isHebrew = language == XMessageLanguage.Hebrew;
                form.RightToLeft = isHebrew ? RightToLeft.Yes : RightToLeft.No;
                form.RightToLeftLayout = isHebrew;

                Label labelPrompt = new Label
                {
                    AutoSize = true,
                    Location = new Point(15, 15),
                    Text = prompt,
                    TextAlign = isHebrew ? ContentAlignment.MiddleRight : ContentAlignment.MiddleLeft,
                    RightToLeft = isHebrew ? RightToLeft.Yes : RightToLeft.No
                };

                TextBox textBox = new TextBox
                {
                    Location = new Point(16, 32),
                    Size = new Size(416, 20),
                    Text = defaultResponse,
                    TextAlign = isHebrew ? HorizontalAlignment.Right : HorizontalAlignment.Left,
                    RightToLeft = isHebrew ? RightToLeft.Yes : RightToLeft.No
                };

                string okText = isHebrew ? "אישור" : "OK";
                string cancelText = isHebrew ? "ביטול" : "Cancel";

                Button buttonOK = new Button
                {
                    DialogResult = DialogResult.OK,
                    Location = new Point(288, 72),
                    Size = new Size(75, 23),
                    Text = okText,
                    TextAlign = ContentAlignment.MiddleCenter,
                    RightToLeft = isHebrew ? RightToLeft.Yes : RightToLeft.No
                };

                Button buttonCancel = new Button
                {
                    DialogResult = DialogResult.Cancel,
                    Location = new Point(376, 72),
                    Size = new Size(75, 23),
                    Text = cancelText,
                    TextAlign = ContentAlignment.MiddleCenter,
                    RightToLeft = isHebrew ? RightToLeft.Yes : RightToLeft.No
                };

                ErrorProvider errorProvider = new ErrorProvider();
                errorProvider.ContainerControl = form;

                if (validator != null)
                {
                    textBox.TextChanged += (s, e) => errorProvider.SetError(textBox, "");
                    textBox.Validating += (s, e) =>
                    {
                        var args = new InputBoxValidatingArgs
                        {
                            Text = textBox.Text
                        };
                        validator(form, args);
                        if (args.Cancel)
                        {
                            e.Cancel = true;
                            errorProvider.SetError(textBox, args.Message);
                        }
                    };
                }

                form.Controls.AddRange(new Control[] { labelPrompt, textBox, buttonOK, buttonCancel });
                form.AcceptButton = buttonOK;
                form.CancelButton = buttonCancel;

                InputBoxResult result = new InputBoxResult();
                if (form.ShowDialog() == DialogResult.OK)
                {
                    result.Text = textBox.Text;
                    result.OK = true;
                }
                else
                {
                    result.Text = textBox.Text;
                    result.OK = false;
                }

                return result;
            }
        }
    }
}

// Usage example:
// var result = XMessageBox.Show("האם אתה בטוח שברצונך לצאת?", "אישור יציאה", XMessageBoxButtons.YesNo, XMessageBoxIcon.Question, XMessageAlignment.Right, XMessageLanguage.Hebrew);
// var inputResult = XMessageBox.ShowInput("הכנס שם קובץ:", "קלט", "", null, -1, -1, XMessageLanguage.Hebrew);






