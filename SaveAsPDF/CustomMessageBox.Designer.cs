using System.Drawing;
using System.Windows.Forms;
using System;

namespace SaveAsPDF
{

    public class CustomMessageBoxForm : Form
    {
        private Label messageLabel;
        private PictureBox iconBox;
        private FlowLayoutPanel buttonPanel;
        private CustomMessageLanguage language;

        public CustomMessageBoxForm(string text, string caption, CustomMessageBoxButtons buttons, CustomMessageBoxIcon icon, CustomMessageAlignment alignment, CustomMessageLanguage language)
        {
            this.language = language;
            this.Text = caption;
            this.RightToLeftLayout = true;
            this.RightToLeft = alignment == CustomMessageAlignment.Right ? RightToLeft.Yes : RightToLeft.No;
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
            messageLabel.TextAlign = GetAlignment(alignment);
            messageLabel.RightToLeft = alignment == CustomMessageAlignment.Right ? RightToLeft.Yes : RightToLeft.No;
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
            buttonPanel.FlowDirection = alignment == CustomMessageAlignment.Right ? FlowDirection.RightToLeft : FlowDirection.LeftToRight;

            AddButtons(buttons, alignment);

            if (alignment == CustomMessageAlignment.Right)
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

        private ContentAlignment GetAlignment(CustomMessageAlignment alignment)
        {
            switch (alignment)
            {
                case CustomMessageAlignment.Left: return ContentAlignment.TopLeft;
                case CustomMessageAlignment.Center: return ContentAlignment.TopCenter;
                default: return ContentAlignment.TopRight;
            }
        }

        private void AddButtons(CustomMessageBoxButtons buttons, CustomMessageAlignment alignment)
        {
            Func<string, string> L = delegate (string key)
            {
                if (language == CustomMessageLanguage.English)
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

            Action<string, DialogResult> Add = delegate (string text, DialogResult result)
            {
                var btn = new Button();
                btn.Text = text;
                btn.DialogResult = result;
                btn.AutoSize = true;
                btn.Click += delegate { this.DialogResult = result; this.Close(); };
                buttonPanel.Controls.Add(btn);
            };

            switch (buttons)
            {
                case CustomMessageBoxButtons.OK:
                    Add(L("OK"), DialogResult.OK);
                    break;
                case CustomMessageBoxButtons.OKCancel:
                    Add(L("Cancel"), DialogResult.Cancel);
                    Add(L("OK"), DialogResult.OK);
                    break;
                case CustomMessageBoxButtons.YesNo:
                    Add(L("No"), DialogResult.No);
                    Add(L("Yes"), DialogResult.Yes);
                    break;
                case CustomMessageBoxButtons.YesNoCancel:
                    Add(L("Cancel"), DialogResult.Cancel);
                    Add(L("No"), DialogResult.No);
                    Add(L("Yes"), DialogResult.Yes);
                    break;
                case CustomMessageBoxButtons.RetryCancel:
                    Add(L("Cancel"), DialogResult.Cancel);
                    Add(L("Retry"), DialogResult.Retry);
                    break;
            }
        }

        private Image GetIconImage(CustomMessageBoxIcon icon)
        {
            switch (icon)
            {
                case CustomMessageBoxIcon.Information: return SystemIcons.Information.ToBitmap();
                case CustomMessageBoxIcon.Warning: return SystemIcons.Warning.ToBitmap();
                case CustomMessageBoxIcon.Error: return SystemIcons.Error.ToBitmap();
                case CustomMessageBoxIcon.Question: return SystemIcons.Question.ToBitmap();
                default: return null;
            }
        }
    }
}