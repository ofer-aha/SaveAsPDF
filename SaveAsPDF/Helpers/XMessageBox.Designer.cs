using System;
using System.Drawing;
using System.Windows;
using System.Windows.Forms;

namespace SaveAsPDF
{
    /// <summary>
    /// Custom message box form that supports right-to-left layouts and multilingual buttons.
    /// </summary>
    public class XMessageBoxForm : Form
    {
        private Label messageLabel;
        private PictureBox iconBox;
        private FlowLayoutPanel buttonPanel;
        private XMessageLanguage language;

        /// <summary>
        /// Initializes a new instance of the <see cref="XMessageBoxForm"/> class.
        /// </summary>
        /// <param name="text">The message text to display.</param>
        /// <param name="caption">The caption for the message box.</param>
        /// <param name="buttons">The buttons to display in the message box.</param>
        /// <param name="icon">The icon to display in the message box.</param>
        /// <param name="alignment">The text alignment (default: Left).</param>
        /// <param name="language">The language for button text (default: English).</param>
        public XMessageBoxForm(string text, string caption, XMessageBoxButtons buttons, XMessageBoxIcon icon,
                XMessageAlignment alignment = XMessageAlignment.Left,
                XMessageLanguage language = XMessageLanguage.English)
        {
            this.language = language;
            this.Text = caption;
            this.RightToLeftLayout = alignment == XMessageAlignment.Right;
            this.RightToLeft = alignment == XMessageAlignment.Right ? RightToLeft.Yes : RightToLeft.No;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.StartPosition = FormStartPosition.CenterScreen;
            this.MinimizeBox = false;
            this.MaximizeBox = false;
            this.Size = new Size(420, 200);

            var table = new TableLayoutPanel();
            table.Dock = DockStyle.Fill;
            table.ColumnCount = 2;
            table.RowCount = 2;
            table.RowStyles.Add(new RowStyle(SizeType.Percent, 70));
            table.RowStyles.Add(new RowStyle(SizeType.Percent, 30));
            table.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            table.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 64));

            iconBox = new PictureBox();
            iconBox.Size = new Size(48, 48);
            iconBox.Margin = new Padding(12);
            iconBox.SizeMode = PictureBoxSizeMode.StretchImage;
            iconBox.Image = GetIconImage(icon);

            messageLabel = new Label();
            messageLabel.Text = text;
            messageLabel.Dock = DockStyle.Fill;
            messageLabel.Padding = new Padding(8);
            messageLabel.TextAlign = alignment == XMessageAlignment.Right ? ContentAlignment.MiddleRight :
                                     alignment == XMessageAlignment.Center ? ContentAlignment.MiddleCenter :
                                     ContentAlignment.MiddleLeft;
            messageLabel.RightToLeft = alignment == XMessageAlignment.Right ? RightToLeft.Yes : RightToLeft.No;
            messageLabel.AutoSize = false;

            Panel messagePanel = new Panel();
            messagePanel.Dock = DockStyle.Fill;
            messagePanel.Padding = new Padding(8);
            messagePanel.Margin = new Padding(4);
            messagePanel.Controls.Add(messageLabel);
            messageLabel.Dock = DockStyle.Fill;

            buttonPanel = new FlowLayoutPanel();
            buttonPanel.Dock = DockStyle.Fill;
            buttonPanel.AutoSize = true;
            buttonPanel.FlowDirection = alignment == XMessageAlignment.Right ? FlowDirection.RightToLeft : FlowDirection.LeftToRight;

            AddButtons(buttons, alignment);

            if (alignment == XMessageAlignment.Right)
            {
                table.Controls.Add(iconBox, 1, 0);
                table.SetRowSpan(iconBox, 2);
                table.Controls.Add(messagePanel, 0, 0);
                table.Controls.Add(buttonPanel, 0, 1);
            }
            else
            {
                table.Controls.Add(iconBox, 0, 0);
                table.SetRowSpan(iconBox, 2);
                table.Controls.Add(messagePanel, 1, 0);
                table.Controls.Add(buttonPanel, 1, 1);
            }

            Controls.Add(table);
        }

        /// <summary>
        /// Gets the appropriate ContentAlignment based on the XMessageAlignment.
        /// </summary>
        /// <param name="alignment">The custom alignment.</param>
        /// <returns>The corresponding ContentAlignment.</returns>
        private ContentAlignment GetAlignment(XMessageAlignment alignment)
        {
            switch (alignment)
            {
                case XMessageAlignment.Right: return ContentAlignment.TopRight;
                case XMessageAlignment.Center: return ContentAlignment.TopCenter;
                default: return ContentAlignment.TopLeft;
            }
        }

        /// <summary>
        /// Adds buttons to the message box based on the specified button configuration.
        /// </summary>
        /// <param name="buttons">The button configuration to use.</param>
        /// <param name="alignment">The alignment of the buttons.</param>
        private void AddButtons(XMessageBoxButtons buttons, XMessageAlignment alignment)
        {
            // Function to get the localized text for button labels
            Func<string, string> L = delegate (string key)
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

            // Function to create and add a button
            Action<string, DialogResult> Add = delegate (string text, DialogResult result)
            {
                var btn = new Button();
                btn.Text = text;
                btn.DialogResult = result;
                btn.AutoSize = true;
                btn.Click += delegate { this.DialogResult = result; this.Close(); };
                buttonPanel.Controls.Add(btn);
            };

            // Add buttons based on the button configuration
            switch (buttons)
            {
                case XMessageBoxButtons.OK:
                    Add(L("OK"), DialogResult.OK);
                    break;
                case XMessageBoxButtons.OKCancel:
                    Add(L("Cancel"), DialogResult.Cancel);
                    Add(L("OK"), DialogResult.OK);
                    break;
                case XMessageBoxButtons.YesNo:
                    Add(L("No"), DialogResult.No);
                    Add(L("Yes"), DialogResult.Yes);
                    break;
                case XMessageBoxButtons.YesNoCancel:
                    Add(L("Cancel"), DialogResult.Cancel);
                    Add(L("No"), DialogResult.No);
                    Add(L("Yes"), DialogResult.Yes);
                    break;
                case XMessageBoxButtons.RetryCancel:
                    Add(L("Cancel"), DialogResult.Cancel);
                    Add(L("Retry"), DialogResult.Retry);
                    break;
            }
        }

        /// <summary>
        /// Gets the appropriate system icon image based on the specified icon type.
        /// </summary>
        /// <param name="icon">The type of icon to retrieve.</param>
        /// <returns>The system icon as an image.</returns>
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
