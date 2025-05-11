using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;

namespace SaveAsPDF
{
    public enum XMessageBoxButtons { OK, OKCancel, YesNo, YesNoCancel, RetryCancel }
    public enum XMessageBoxIcon { None, Information, Warning, Error, Question }
    public enum XMessageAlignment { Left, Center, Right }
    public enum XMessageLanguage { Hebrew, English }

    public static class XMessageBox
    {
        public static DialogResult Show(string text, string caption, XMessageBoxButtons buttons, XMessageBoxIcon icon, XMessageAlignment alignment = XMessageAlignment.Right, XMessageLanguage language = XMessageLanguage.Hebrew)
        {
            using (var msgBox = new XMessageBoxForm(text, caption, buttons, icon, alignment, language))
            {
                return msgBox.ShowDialog();
            }
        }
    }

    public class XMessageBoxForm : Form
    {
        private Label messageLabel;
        private TextBox messageTextBox;
        private PictureBox iconBox;
        private FlowLayoutPanel buttonPanel;
        private XMessageLanguage language;
        private Timer autoCloseTimer;

        public XMessageBoxForm(string text, string caption, XMessageBoxButtons buttons, XMessageBoxIcon icon,
            XMessageAlignment alignment = XMessageAlignment.Left,
            XMessageLanguage language = XMessageLanguage.English,
            bool useTextBox = false,
            bool isDarkTheme = false,
            Font customFont = null,
            Color? backColorOverride = null,
            int autoCloseMilliseconds = 0,
            Image customIcon = null)
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
            this.BackColor = backColorOverride ?? (isDarkTheme ? Color.FromArgb(40, 40, 40) : SystemColors.Control);

            TableLayoutPanel table = new TableLayoutPanel();
            table.Dock = DockStyle.Fill;
            table.ColumnCount = 2;
            table.RowCount = 2;
            table.AutoSize = true;

            table.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            table.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            table.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            table.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));

            iconBox = new PictureBox
            {
                Size = new Size(48, 48),
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
                    MaximumSize = new Size(400, 300),
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
                    MaximumSize = new Size(400, 0),
                    AutoEllipsis = true,
                    AccessibleName = "Message",
                    AccessibleDescription = "Message displayed in the custom message box",
                    Font = customFont ?? new Font("Segoe UI", 10),
                    ForeColor = isDarkTheme ? Color.White : SystemColors.ControlText
                };
            }

            buttonPanel = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                AutoSize = true,
                FlowDirection = alignment == XMessageAlignment.Right ? FlowDirection.RightToLeft : FlowDirection.LeftToRight
            };

            AddButtons(buttons);

            // Layout order — fixed, RTL mirrored by WinForms
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
        }

        private ContentAlignment GetContentAlignment(XMessageAlignment alignment)
        {
            switch (alignment)
            {
                case XMessageAlignment.Right: return ContentAlignment.MiddleRight;
                case XMessageAlignment.Center: return ContentAlignment.MiddleCenter;
                default: return ContentAlignment.MiddleLeft;
            }
        }

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
