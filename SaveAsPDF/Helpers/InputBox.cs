using System;
using System.ComponentModel;
using System.Windows.Forms;

namespace SaveAsPDF.Helpers
{
    /// <summary>
    /// The InputBox class is used to show a prompt in a dialog box using the static method Show().
    /// </summary>
    public class InputBox : Form
    {
        private Button buttonOK;
        private Button buttonCancel;
        private Label labelPrompt;
        private TextBox textBoxText;
        private ErrorProvider errorProviderText;
        private IContainer components;
        private InputBoxValidatingHandler _validator;

        /// <summary>
        /// Initializes a new instance of the <see cref="InputBox"/> class.
        /// </summary>
        private InputBox()
        {
            InitializeComponent();
        }

        /// <summary>
        /// Releases the resources used by the <see cref="InputBox"/> class.
        /// </summary>
        /// <param name="disposing">True to release both managed and unmanaged resources; false to release only unmanaged resources.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                components?.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code
        /// <summary>
        /// Initializes the components of the InputBox form.
        /// </summary>
        private void InitializeComponent()
        {
            components = new Container();
            buttonOK = new Button();
            buttonCancel = new Button();
            textBoxText = new TextBox();
            labelPrompt = new Label();
            errorProviderText = new ErrorProvider(components);

            ((ISupportInitialize)(errorProviderText)).BeginInit();
            SuspendLayout();

            // buttonOK
            buttonOK.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            buttonOK.DialogResult = DialogResult.OK;
            buttonOK.Location = new System.Drawing.Point(288, 72);
            buttonOK.Name = "buttonOK";
            buttonOK.Size = new System.Drawing.Size(75, 23);
            buttonOK.TabIndex = 2;
            buttonOK.Text = "אישור";
            buttonOK.Click += buttonOK_Click;

            // buttonCancel
            buttonCancel.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            buttonCancel.CausesValidation = false;
            buttonCancel.DialogResult = DialogResult.Cancel;
            buttonCancel.Location = new System.Drawing.Point(376, 72);
            buttonCancel.Name = "buttonCancel";
            buttonCancel.Size = new System.Drawing.Size(75, 23);
            buttonCancel.TabIndex = 3;
            buttonCancel.Text = "ביטול";
            buttonCancel.Click += buttonCancel_Click;

            // textBoxText
            textBoxText.Location = new System.Drawing.Point(16, 32);
            textBoxText.Name = "textBoxText";
            textBoxText.Size = new System.Drawing.Size(416, 20);
            textBoxText.TabIndex = 1;
            textBoxText.TextChanged += textBoxText_TextChanged;
            textBoxText.Validating += textBoxText_Validating;

            // labelPrompt
            labelPrompt.AutoSize = true;
            labelPrompt.Location = new System.Drawing.Point(15, 15);
            labelPrompt.Name = "labelPrompt";
            labelPrompt.Size = new System.Drawing.Size(39, 13);
            labelPrompt.TabIndex = 0;
            labelPrompt.Text = "prompt";

            // errorProviderText
            errorProviderText.ContainerControl = this;

            // InputBox
            AcceptButton = buttonOK;
            AutoScaleBaseSize = new System.Drawing.Size(5, 13);
            CancelButton = buttonCancel;
            ClientSize = new System.Drawing.Size(464, 104);
            Controls.Add(labelPrompt);
            Controls.Add(textBoxText);
            Controls.Add(buttonCancel);
            Controls.Add(buttonOK);
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            Name = "InputBox";
            RightToLeft = RightToLeft.Yes;
            RightToLeftLayout = true;
            Text = "Title";

            ((ISupportInitialize)(errorProviderText)).EndInit();
            ResumeLayout(false);
            PerformLayout();
        }
        #endregion

        /// <summary>
        /// Handles the Cancel button click event.
        /// </summary>
        private void buttonCancel_Click(object sender, EventArgs e)
        {
            Validator = null;
            Close();
        }

        /// <summary>
        /// Handles the OK button click event.
        /// </summary>
        private void buttonOK_Click(object sender, EventArgs e)
        {
            Close();
        }

        /// <summary>
        /// Displays a prompt in a dialog box, waits for the user to input text or click a button.
        /// </summary>
        /// <param name="prompt">The message to display in the dialog box.</param>
        /// <param name="title">The title of the dialog box.</param>
        /// <param name="defaultResponse">The default text to display in the input box.</param>
        /// <param name="validator">A delegate to validate the input text.</param>
        /// <param name="xpos">The X-coordinate of the dialog box on the screen.</param>
        /// <param name="ypos">The Y-coordinate of the dialog box on the screen.</param>
        /// <returns>An <see cref="InputBoxResult"/> containing the input text and whether the OK button was clicked.</returns>
        public static InputBoxResult Show(string prompt, string title, string defaultResponse = "", InputBoxValidatingHandler validator = null, int xpos = -1, int ypos = -1)
        {
            using (InputBox form = new InputBox())
            {
                form.labelPrompt.Text = prompt;
                form.Text = title;
                form.textBoxText.Text = defaultResponse;

                if (xpos >= 0 && ypos >= 0)
                {
                    form.StartPosition = FormStartPosition.Manual;
                    form.Left = xpos;
                    form.Top = ypos;
                }

                form.Validator = validator;

                DialogResult result = form.ShowDialog();

                return new InputBoxResult
                {
                    Text = form.textBoxText.Text,
                    OK = result == DialogResult.OK
                };
            }
        }

        /// <summary>
        /// Resets the error provider when the text in the input box changes.
        /// </summary>
        private void textBoxText_TextChanged(object sender, EventArgs e)
        {
            errorProviderText.SetError(textBoxText, "");
        }

        /// <summary>
        /// Validates the input text using the provided validator.
        /// </summary>
        private void textBoxText_Validating(object sender, CancelEventArgs e)
        {
            if (Validator != null)
            {
                var args = new InputBoxValidatingArgs
                {
                    Text = textBoxText.Text
                };
                Validator(this, args);

                if (args.Cancel)
                {
                    e.Cancel = true;
                    errorProviderText.SetError(textBoxText, args.Message);
                }
            }
        }

        /// <summary>
        /// Gets or sets the validator delegate for the input box.
        /// </summary>
        protected InputBoxValidatingHandler Validator
        {
            get => _validator;
            set => _validator = value;
        }
    }

    /// <summary>
    /// Class used to store the result of an InputBox.Show message.
    /// </summary>
    public class InputBoxResult
    {
        /// <summary>
        /// Gets or sets a value indicating whether the OK button was clicked.
        /// </summary>
        public bool OK { get; set; }

        /// <summary>
        /// Gets or sets the text entered in the input box.
        /// </summary>
        public string Text { get; set; }
    }

    /// <summary>
    /// EventArgs used to validate an InputBox.
    /// </summary>
    public class InputBoxValidatingArgs : EventArgs
    {
        /// <summary>
        /// Gets or sets the text entered in the input box.
        /// </summary>
        public string Text { get; set; }

        /// <summary>
        /// Gets or sets the validation error message.
        /// </summary>
        public string Message { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether the validation failed.
        /// </summary>
        public bool Cancel { get; set; }
    }

    /// <summary>
    /// Delegate used to validate an InputBox.
    /// </summary>
    /// <param name="sender">The source of the event.</param>
    /// <param name="e">The <see cref="InputBoxValidatingArgs"/> containing the event data.</param>
    public delegate void InputBoxValidatingHandler(object sender, InputBoxValidatingArgs e);

}
