using System;
using System.Windows.Forms;
using System.ComponentModel;

namespace SaveAsPDF
{
    public enum XMessageBoxButtons { OK, OKCancel, YesNo, YesNoCancel, RetryCancel }
    public enum XMessageBoxIcon { None, Information, Warning, Error, Question }
    public enum XMessageAlignment { Left, Center, Right }
    public enum XMessageLanguage { Hebrew, English }

    public class InputBoxResult
    {
        public bool OK { get; set; }
        public string Text { get; set; }
    }

    public class InputBoxValidatingArgs : EventArgs
    {
        public string Text { get; set; }
        public string Message { get; set; }
        public bool Cancel { get; set; }
    }

    public delegate void InputBoxValidatingHandler(object sender, InputBoxValidatingArgs e);

    public static class XMessageBox
    {
        public static DialogResult Show(string text, string caption, XMessageBoxButtons buttons, XMessageBoxIcon icon, XMessageAlignment alignment = XMessageAlignment.Right, XMessageLanguage language = XMessageLanguage.Hebrew)
        {
            using (var msgBox = new XMessageBoxForm(text, caption, buttons, icon, alignment, language))
            {
                return msgBox.ShowDialog();
            }
        }

        public static InputBoxResult ShowInput(string prompt, string title, string defaultResponse = "", InputBoxValidatingHandler validator = null, int xpos = -1, int ypos = -1)
        {
            using (Form form = new Form())
            {
                form.Text = title;
                form.FormBorderStyle = FormBorderStyle.FixedDialog;
                form.MaximizeBox = false;
                form.MinimizeBox = false;
                form.RightToLeft = RightToLeft.Yes;
                form.RightToLeftLayout = true;
                form.ClientSize = new System.Drawing.Size(464, 104);
                form.StartPosition = (xpos >= 0 && ypos >= 0) ? FormStartPosition.Manual : FormStartPosition.CenterScreen;
                if (xpos >= 0 && ypos >= 0)
                {
                    form.Left = xpos;
                    form.Top = ypos;
                }

                Label labelPrompt = new Label
                {
                    AutoSize = true,
                    Location = new System.Drawing.Point(15, 15),
                    Text = prompt
                };

                TextBox textBox = new TextBox
                {
                    Location = new System.Drawing.Point(16, 32),
                    Size = new System.Drawing.Size(416, 20),
                    Text = defaultResponse
                };

                Button buttonOK = new Button
                {
                    DialogResult = DialogResult.OK,
                    Location = new System.Drawing.Point(288, 72),
                    Size = new System.Drawing.Size(75, 23),
                    Text = "אישור"
                };

                Button buttonCancel = new Button
                {
                    DialogResult = DialogResult.Cancel,
                    Location = new System.Drawing.Point(376, 72),
                    Size = new System.Drawing.Size(75, 23),
                    Text = "ביטול"
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

                return result;
            }
        }
    }
}


// Usage example:
// var result = XMessageBox.Show("האם אתה בטוח שברצונך לצאת?", "אישור יציאה", XMessageBoxButtons.YesNo, XMessageBoxIcon.Question);






