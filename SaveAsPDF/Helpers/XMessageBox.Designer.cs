using System;
using System.Drawing;
using System.Windows.Forms;

namespace SaveAsPDF
{
    /// <summary>
    /// Custom message box form that supports right-to-left layouts, multilingual buttons, and auto-close timer in the title.
    /// </summary>
    public class XMessageBoxForm : Form
    {
        private Label messageLabel;
        private PictureBox iconBox;
        private FlowLayoutPanel buttonPanel;
        private XMessageLanguage language;
        private Timer autoCloseTimer;
        private int remainingSeconds;
        private int autoCloseMilliseconds;
        private string baseTitle;

        /// <summary>
        /// Initializes a new instance of the <see cref="XMessageBoxForm"/> class.
        /// </summary>
        public XMessageBoxForm(
            string text,
            string caption,
            XMessageBoxButtons buttons,
            XMessageBoxIcon icon,
            XMessageAlignment alignment = XMessageAlignment.Left,
            XMessageLanguage language = XMessageLanguage.English,
            int autoCloseMilliseconds = 0)
        {
            this.language = language;
            this.baseTitle = caption ?? string.Empty;
            this.Text = baseTitle;
            bool isHebrew = language == XMessageLanguage.Hebrew;
            this.RightToLeftLayout = isHebrew || alignment == XMessageAlignment.Right;
            this.RightToLeft = isHebrew || alignment == XMessageAlignment.Right ? RightToLeft.Yes : RightToLeft.No;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.StartPosition = FormStartPosition.CenterScreen;
            this.MinimizeBox = false;
            this.MaximizeBox = false;
            this.Size = new System.Drawing.Size(420, 200);
            this.autoCloseMilliseconds = autoCloseMilliseconds;

            var table = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 2,
                RowCount = 2
            };
            table.RowStyles.Add(new RowStyle(SizeType.Percent, 70));
            table.RowStyles.Add(new RowStyle(SizeType.Percent, 30));
            table.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 64));
            table.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            table.RightToLeft = isHebrew ? RightToLeft.Yes : RightToLeft.No;

            iconBox = new PictureBox
            {
                Size = new System.Drawing.Size(48, 48),
                Margin = new Padding(12),
                SizeMode = PictureBoxSizeMode.StretchImage,
                Image = GetIconImage(icon),
                RightToLeft = isHebrew ? RightToLeft.Yes : RightToLeft.No
            };

            messageLabel = new Label
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(8, 8, 8, 2),
                AutoSize = false,
                Font = new Font(System.Drawing.SystemFonts.MessageBoxFont.FontFamily, System.Drawing.SystemFonts.MessageBoxFont.Size, System.Drawing.SystemFonts.MessageBoxFont.Style),
                Text = text
            };

            // Hebrew labels always align right, otherwise use alignment
            if (isHebrew)
            {
                messageLabel.TextAlign = ContentAlignment.MiddleRight;
                messageLabel.RightToLeft = RightToLeft.Yes;
            }
            else
            {
                switch (alignment)
                {
                    case XMessageAlignment.Right:
                        messageLabel.TextAlign = ContentAlignment.MiddleRight;
                        messageLabel.RightToLeft = RightToLeft.Yes;
                        break;
                    case XMessageAlignment.Center:
                        messageLabel.TextAlign = ContentAlignment.MiddleCenter;
                        messageLabel.RightToLeft = RightToLeft.No;
                        break;
                    default:
                        messageLabel.TextAlign = ContentAlignment.MiddleLeft;
                        messageLabel.RightToLeft = RightToLeft.No;
                        break;
                }
            }

            buttonPanel = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                AutoSize = true,
                FlowDirection = (isHebrew || alignment == XMessageAlignment.Right)
                    ? System.Windows.Forms.FlowDirection.RightToLeft
                    : System.Windows.Forms.FlowDirection.LeftToRight,
                RightToLeft = isHebrew ? RightToLeft.Yes : (alignment == XMessageAlignment.Right ? RightToLeft.Yes : RightToLeft.No)
            };

            AddButtons(buttons, alignment);

            // Layout: icon left, message right (or reversed for RTL)
            if (isHebrew || alignment == XMessageAlignment.Right)
            {
                table.Controls.Add(iconBox, 1, 0);
                table.SetRowSpan(iconBox, 2);
                table.Controls.Add(messageLabel, 0, 0);
                table.Controls.Add(buttonPanel, 0, 1);
            }
            else
            {
                table.Controls.Add(iconBox, 0, 0);
                table.SetRowSpan(iconBox, 2);
                table.Controls.Add(messageLabel, 1, 0);
                table.Controls.Add(buttonPanel, 1, 1);
            }

            Controls.Add(table);

            // Setup auto-close timer if needed
            if (autoCloseMilliseconds > 0)
            {
                remainingSeconds = (int)Math.Ceiling(autoCloseMilliseconds / 1000.0);
                UpdateTitleWithTimer(remainingSeconds);
                autoCloseTimer = new Timer();
                autoCloseTimer.Interval = 1000;
                autoCloseTimer.Tick += (s, e) =>
                {
                    remainingSeconds--;
                    UpdateTitleWithTimer(remainingSeconds);
                    if (remainingSeconds <= 0)
                    {
                        autoCloseTimer.Stop();
                        // Find the default button and perform click
                        Button defaultButton = null;
                        foreach (Control c in buttonPanel.Controls)
                        {
                            if (c is Button btn && btn.DialogResult == GetDefaultDialogResult(buttons))
                            {
                                defaultButton = btn;
                                break;
                            }
                        }
                        if (defaultButton != null)
                        {
                            defaultButton.PerformClick();
                        }
                        else
                        {
                            this.DialogResult = GetDefaultDialogResult(buttons);
                            this.Close();
                        }
                    }
                };
                autoCloseTimer.Start();
            }
        }

        /// <summary>
        /// Updates the message box title to include the auto-close timer in the appropriate language.
        /// </summary>
        /// <param name="seconds">The number of seconds remaining before auto-close.</param>
        private void UpdateTitleWithTimer(int seconds)
        {
            string timerMsg;
            if (language == XMessageLanguage.Hebrew)
                timerMsg = $" (סוגר בעוד: {seconds} {(seconds == 1 ? "שנייה" : "שניות")})";
            else
                timerMsg = $" (Closing in: {seconds} second{(seconds == 1 ? "" : "s")})";
            this.Text = baseTitle + timerMsg;
        }

        /// <summary>
        /// Gets the default <see cref="DialogResult"/> for the specified button configuration.
        /// </summary>
        private DialogResult GetDefaultDialogResult(XMessageBoxButtons buttons)
        {
            switch (buttons)
            {
                case XMessageBoxButtons.OKCancel:
                    return XMessageBox.DefaultDialogResultForOKCancel(language);
                case XMessageBoxButtons.YesNo:
                    return XMessageBox.DefaultDialogResultForYesNo(language);
                case XMessageBoxButtons.YesNoCancel:
                    return XMessageBox.DefaultDialogResultForYesNoCancel(language);
                case XMessageBoxButtons.RetryCancel:
                    return XMessageBox.DefaultDialogResultForRetryCancel(language);
                default:
                    return DialogResult.OK;
            }
        }

        /// <summary>
        /// Adds buttons to the message box based on the specified button configuration and alignment.
        /// </summary>
        private void AddButtons(XMessageBoxButtons buttons, XMessageAlignment alignment)
        {
            Func<string, string> L = key =>
            {
                if (language == XMessageLanguage.English)
                {
                    switch (key)
                    {
                        case "OK": return "OK";
                        case "Cancel": return "Cancel";
                        case "Yes": return "Yes";
                        case "No": return "No";
                        case "Retry": return "Retry";
                    }
                }
                else
                {
                    switch (key)
                    {
                        case "OK": return "אישור";
                        case "Cancel": return "ביטול";
                        case "Yes": return "כן";
                        case "No": return "לא";
                        case "Retry": return "נסה שוב";
                    }
                }
                return key;
            };

            Action<string, DialogResult> Add = (text, result) =>
            {
                var btn = new Button
                {
                    Text = text,
                    DialogResult = result,
                    AutoSize = true,
                    TextAlign = ContentAlignment.MiddleCenter,
                    RightToLeft = language == XMessageLanguage.Hebrew ? RightToLeft.Yes : RightToLeft.No
                };
                btn.Click += (s, e) => { this.DialogResult = result; this.Close(); };
                buttonPanel.Controls.Add(btn);
            };

            // Add buttons in correct order for RTL/LTR
            switch (buttons)
            {
                case XMessageBoxButtons.OK:
                    Add(L("OK"), DialogResult.OK);
                    break;
                case XMessageBoxButtons.OKCancel:
                    if (language == XMessageLanguage.Hebrew)
                    {
                        Add(L("Cancel"), DialogResult.Cancel);
                        Add(L("OK"), DialogResult.OK);
                    }
                    else
                    {
                        Add(L("OK"), DialogResult.OK);
                        Add(L("Cancel"), DialogResult.Cancel);
                    }
                    break;
                case XMessageBoxButtons.YesNo:
                    if (language == XMessageLanguage.Hebrew)
                    {
                        Add(L("No"), DialogResult.No);
                        Add(L("Yes"), DialogResult.Yes);
                    }
                    else
                    {
                        Add(L("Yes"), DialogResult.Yes);
                        Add(L("No"), DialogResult.No);
                    }
                    break;
                case XMessageBoxButtons.YesNoCancel:
                    if (language == XMessageLanguage.Hebrew)
                    {
                        Add(L("Cancel"), DialogResult.Cancel);
                        Add(L("No"), DialogResult.No);
                        Add(L("Yes"), DialogResult.Yes);
                    }
                    else
                    {
                        Add(L("Yes"), DialogResult.Yes);
                        Add(L("No"), DialogResult.No);
                        Add(L("Cancel"), DialogResult.Cancel);
                    }
                    break;
                case XMessageBoxButtons.RetryCancel:
                    if (language == XMessageLanguage.Hebrew)
                    {
                        Add(L("Cancel"), DialogResult.Cancel);
                        Add(L("Retry"), DialogResult.Retry);
                    }
                    else
                    {
                        Add(L("Retry"), DialogResult.Retry);
                        Add(L("Cancel"), DialogResult.Cancel);
                    }
                    break;
            }
        }

        /// <summary>
        /// Gets the appropriate system icon image based on the specified icon type.
        /// </summary>
        private Image GetIconImage(XMessageBoxIcon icon)
        {
            switch (icon)
            {
                case XMessageBoxIcon.Information: return SystemIcons.Information.ToBitmap();
                case XMessageBoxIcon.Warning: return SystemIcons.Warning.ToBitmap();
                case XMessageBoxIcon.Error: return SystemIcons.Error.ToBitmap();
                case XMessageBoxIcon.Question: return SystemIcons.Question.ToBitmap();
                default: return null;
            }
        }

        /// <summary>
        /// Releases the resources used by the <see cref="XMessageBoxForm"/>.
        /// </summary>
        protected override void Dispose(bool disposing)
        {
            if (disposing && autoCloseTimer != null)
            {
                autoCloseTimer.Dispose();
            }
            base.Dispose(disposing);
        }
    }

    /// <summary>
    /// Helper methods for default dialog results for RTL/LTR.
    /// </summary>
    public static partial class XMessageBox
    {
        public static DialogResult DefaultDialogResultForOKCancel(XMessageLanguage lang)
        {
            return lang == XMessageLanguage.Hebrew ? DialogResult.OK : DialogResult.OK;
        }
        public static DialogResult DefaultDialogResultForYesNo(XMessageLanguage lang)
        {
            return lang == XMessageLanguage.Hebrew ? DialogResult.Yes : DialogResult.Yes;
        }
        public static DialogResult DefaultDialogResultForYesNoCancel(XMessageLanguage lang)
        {
            return lang == XMessageLanguage.Hebrew ? DialogResult.Yes : DialogResult.Yes;
        }
        public static DialogResult DefaultDialogResultForRetryCancel(XMessageLanguage lang)
        {
            return lang == XMessageLanguage.Hebrew ? DialogResult.Retry : DialogResult.Retry;
        }
    }
}
