using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows;
using System.Windows.Forms;

namespace SaveAsPDF
{
    /// <summary>
    /// Represents the buttons available in the custom message box.
    /// </summary>
    public enum XMessageBoxButtons { OK, OKCancel, YesNo, YesNoCancel, RetryCancel }

    /// <summary>
    /// Represents the icons available in the custom message box.
    /// </summary>
    public enum XMessageBoxIcon { None, Information, Warning, Error, Question }

    /// <summary>
    /// Represents the alignment options for the message text in the custom message box.
    /// </summary>
    public enum XMessageAlignment { Left, Center, Right }

    /// <summary>
    /// Represents the language options for the custom message box.
    /// </summary>
    public enum XMessageLanguage { Hebrew, English }

    /// <summary>
    /// A custom message box form that provides advanced configuration options.
    /// </summary>
    public class XMessageBoxForm : Form
    {
        private Label messageLabel;
        private TextBox messageTextBox;
        private PictureBox iconBox;
        private FlowLayoutPanel buttonPanel;
        private XMessageLanguage language;
        private Timer autoCloseTimer;

        /// <summary>
        /// Initializes a new instance of the <see cref="XMessageBoxForm"/> class.
        /// </summary>
        /// <param name="text">The message text to display in the message box.</param>
        /// <param name="caption">The caption text for the message box title bar.</param>
        /// <param name="buttons">The buttons to display in the message box.</param>
        /// <param name="icon">The icon to display in the message box.</param>
        /// <param name="alignment">The alignment of the message text.</param>
        /// <param name="language">The language of the message box buttons.</param>
        /// <param name="useTextBox">Indicates whether to display the message in a multi-line text box.</param>
        /// <param name="isDarkTheme">Indicates whether to use a dark theme for the message box.</param>
        /// <param name="customFont">A custom font to use for the message text.</param>
        /// <param name="backColorOverride">An optional background color override for the message box.</param>
        /// <param name="autoCloseMilliseconds">The time in milliseconds after which the message box will automatically close.</param>
        /// <param name="customIcon">An optional custom icon to display in the message box.</param>
        /// <param name="centerToActiveWindow">Indicates whether to center the message box to the active window.</param>
        public XMessageBoxForm(
            string text,
            string caption,
            XMessageBoxButtons buttons,
            XMessageBoxIcon icon,
            XMessageAlignment alignment = XMessageAlignment.Left,
            XMessageLanguage language = XMessageLanguage.English,
            bool useTextBox = false,
            bool isDarkTheme = false,
            Font customFont = null,
            Color? backColorOverride = null,
            int autoCloseMilliseconds = 0,
            Image customIcon = null,
            bool centerToActiveWindow = true)
        {
            this.language = language;
            this.Text = caption;

            this.RightToLeftLayout = alignment == XMessageAlignment.Right;
            this.RightToLeft = alignment == XMessageAlignment.Right ? RightToLeft.Yes : RightToLeft.No;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.StartPosition = FormStartPosition.CenterScreen;
            this.MinimizeBox = false;
            this.MaximizeBox = false;
            this.AutoSize = true;
            this.AutoSizeMode = AutoSizeMode.GrowAndShrink;
            this.BackColor = backColorOverride ?? (isDarkTheme ? Color.FromArgb(40, 40, 40) : System.Drawing.SystemColors.Control);

            TableLayoutPanel table = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 2,
                RowCount = 2,
                AutoSize = true
            };

            table.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            table.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            table.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            table.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));

            iconBox = new PictureBox
            {
                Size = new System.Drawing.Size(48, 48),
                Margin = new Padding(12),
                SizeMode = PictureBoxSizeMode.StretchImage,
                Image = customIcon ?? GetIconImage(icon),
                AccessibleName = "Icon"
            };

            if (useTextBox)
            {
                messageTextBox = new TextBox
                {
                    Multiline = true,
                    ReadOnly = true,
                    BorderStyle = BorderStyle.None,
                    Text = text,
                    BackColor = this.BackColor,
                    Font = customFont ?? new Font("Segoe UI", 10),
                    MaximumSize = new System.Drawing.Size(400, 300),
                    Dock = DockStyle.Fill,
                    ScrollBars = ScrollBars.Vertical
                };
            }
            else
            {
                messageLabel = new Label
                {
                    Text = text,
                    AutoSize = true,
                    Padding = new Padding(8),
                    TextAlign = GetContentAlignment(alignment),
                    RightToLeft = alignment == XMessageAlignment.Right ? RightToLeft.Yes : RightToLeft.No,
                    MaximumSize = new System.Drawing.Size(400, 0),
                    AutoEllipsis = true,
                    AccessibleName = "Message",
                    AccessibleDescription = "Message displayed in the custom message box",
                    Font = customFont ?? new Font("Segoe UI", 10),
                    ForeColor = isDarkTheme ? Color.White : System.Drawing.SystemColors.ControlText
                };
            }

            buttonPanel = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                AutoSize = true,
                FlowDirection = (System.Windows.Forms.FlowDirection)(alignment == XMessageAlignment.Right ? System.Windows.FlowDirection.RightToLeft : System.Windows.FlowDirection.LeftToRight)
            };

            AddButtons(buttons);

            // Layout
            table.Controls.Add(iconBox, 0, 0);
            if (useTextBox)
                table.Controls.Add(messageTextBox, 1, 0);
            else
                table.Controls.Add(messageLabel, 1, 0);

            table.SetRowSpan(iconBox, 2);
            table.Controls.Add(buttonPanel, 1, 1);

            Controls.Add(table);

            // Auto-close
            if (autoCloseMilliseconds > 0)
            {
                autoCloseTimer = new Timer();
                autoCloseTimer.Interval = autoCloseMilliseconds;
                autoCloseTimer.Tick += (s, e) =>
                {
                    autoCloseTimer.Stop();
                    this.DialogResult = DialogResult.OK;
                    this.Close();
                };
                autoCloseTimer.Start();
            }

            // Center to active window
            if (centerToActiveWindow)
            {
                var active = Form.ActiveForm;
                if (active != null)
                {
                    this.StartPosition = FormStartPosition.Manual;
                    this.Location = new System.Drawing.Point(
                        active.Location.X + (active.Width - this.Width) / 2,
                        active.Location.Y + (active.Height - this.Height) / 2
                    );
                }
                else
                {
                    this.StartPosition = FormStartPosition.CenterScreen;
                }
            }
        }

        /// <summary>
        /// Gets the content alignment based on the specified <see cref="XMessageAlignment"/>.
        /// </summary>
        /// <param name="alignment">The alignment option.</param>
        /// <returns>The corresponding <see cref="ContentAlignment"/>.</returns>
        private ContentAlignment GetContentAlignment(XMessageAlignment alignment)
        {
            switch (alignment)
            {
                case XMessageAlignment.Right: return ContentAlignment.MiddleRight;
                case XMessageAlignment.Center: return ContentAlignment.MiddleCenter;
                default: return ContentAlignment.MiddleLeft;
            }
        }

        /// <summary>
        /// Adds buttons to the message box based on the specified <see cref="XMessageBoxButtons"/>.
        /// </summary>
        /// <param name="buttons">The button configuration.</param>
        private void AddButtons(XMessageBoxButtons buttons)
        {
            Dictionary<string, string> localizedTexts = (language == XMessageLanguage.English)
                ? new Dictionary<string, string>
                {
                    ["OK"] = "OK",
                    ["Cancel"] = "Cancel",
                    ["Yes"] = "Yes",
                    ["No"] = "No",
                    ["Retry"] = "Retry"
                }
                : new Dictionary<string, string>
                {
                    ["OK"] = "אישור",
                    ["Cancel"] = "ביטול",
                    ["Yes"] = "כן",
                    ["No"] = "לא",
                    ["Retry"] = "נסה שוב"
                };

            Action<string, DialogResult> AddButton = (key, result) =>
            {
                Button btn = new Button
                {
                    Text = localizedTexts[key],
                    DialogResult = result,
                    AutoSize = true,
                    AccessibleName = key,
                    TabIndex = buttonPanel.Controls.Count
                };
                btn.Click += (s, e) => { this.DialogResult = result; this.Close(); };
                buttonPanel.Controls.Add(btn);

                if (result == DialogResult.OK || result == DialogResult.Yes)
                    this.AcceptButton = btn;
                if (result == DialogResult.Cancel)
                    this.CancelButton = btn;
            };

            switch (buttons)
            {
                case XMessageBoxButtons.OK:
                    AddButton("OK", DialogResult.OK);
                    break;
                case XMessageBoxButtons.OKCancel:
                    AddButton("Cancel", DialogResult.Cancel);
                    AddButton("OK", DialogResult.OK);
                    break;
                case XMessageBoxButtons.YesNo:
                    AddButton("No", DialogResult.No);
                    AddButton("Yes", DialogResult.Yes);
                    break;
                case XMessageBoxButtons.YesNoCancel:
                    AddButton("Cancel", DialogResult.Cancel);
                    AddButton("No", DialogResult.No);
                    AddButton("Yes", DialogResult.Yes);
                    break;
                case XMessageBoxButtons.RetryCancel:
                    AddButton("Cancel", DialogResult.Cancel);
                    AddButton("Retry", DialogResult.Retry);
                    break;
            }
        }

        /// <summary>
        /// Gets the icon image based on the specified <see cref="XMessageBoxIcon"/>.
        /// </summary>
        /// <param name="icon">The icon type.</param>
        /// <returns>The corresponding <see cref="Image"/>.</returns>
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
    }
}
