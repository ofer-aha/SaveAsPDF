﻿using System;
using System.ComponentModel;
using System.Windows.Forms;

namespace SaveAsPDF.Helpers
{
    /// <summary>
    /// The InputBox class is used to show a prompt in a dialog box using the static method Show().
    /// </summary>
    /// <remarks>
    /// Copyright © 2003 Reflection IT
    /// 
    /// This software is provided 'as-is', without any express or implied warranty.
    /// In no event will the authors be held liable for any damages arising from the
    /// use of this software.
    /// 
    /// Permission is granted to anyone to use this software for any purpose,
    /// including commercial applications, subject to the following restrictions:
    /// 
    /// 1. The origin of this software must not be misrepresented; you must not claim
    /// that you wrote the original software. 
    /// 
    /// 2. No substantial portion of the source code of this library may be redistributed
    /// without the express written permission of the copyright holders, where
    /// "substantial" is defined as enough code to be recognizably from this library. 
    /// 
    /// </remarks>
    public class InputBox : System.Windows.Forms.Form
    {
        protected System.Windows.Forms.Button buttonOK;
        protected System.Windows.Forms.Button buttonCancel;
        protected System.Windows.Forms.Label labelPrompt;
        protected System.Windows.Forms.TextBox textBoxText;
        protected System.Windows.Forms.ErrorProvider errorProviderText;
        private IContainer components;

        /// <summary>
        /// Delegate used to validate the object
        /// </summary>
        private InputBoxValidatingHandler _validator;

        private InputBox()
        {
            //
            // Required for Windows Form Designer support
            //
            InitializeComponent();

            //
            // Add any constructor code after InitializeComponent call
            //
        }

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                if (components != null)
                {
                    components.Dispose();
                }
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code
        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            components = new System.ComponentModel.Container();
            buttonOK = new System.Windows.Forms.Button();
            buttonCancel = new System.Windows.Forms.Button();
            textBoxText = new System.Windows.Forms.TextBox();
            labelPrompt = new System.Windows.Forms.Label();
            errorProviderText = new System.Windows.Forms.ErrorProvider(components);
            ((System.ComponentModel.ISupportInitialize)(errorProviderText)).BeginInit();
            SuspendLayout();
            // 
            // buttonOK
            // 
            buttonOK.Anchor = (System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right);
            buttonOK.DialogResult = System.Windows.Forms.DialogResult.OK;
            buttonOK.Location = new System.Drawing.Point(288, 72);
            buttonOK.Name = "buttonOK";
            buttonOK.Size = new System.Drawing.Size(75, 23);
            buttonOK.TabIndex = 2;
            buttonOK.Text = "אישור";
            buttonOK.Click += new System.EventHandler(buttonOK_Click);
            // 
            // buttonCancel
            // 
            buttonCancel.Anchor = (System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right);
            buttonCancel.CausesValidation = false;
            buttonCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            buttonCancel.Location = new System.Drawing.Point(376, 72);
            buttonCancel.Name = "buttonCancel";
            buttonCancel.Size = new System.Drawing.Size(75, 23);
            buttonCancel.TabIndex = 3;
            buttonCancel.Text = "ביטול";
            buttonCancel.Click += new System.EventHandler(buttonCancel_Click);
            // 
            // textBoxText
            // 
            textBoxText.Location = new System.Drawing.Point(16, 32);
            textBoxText.Name = "textBoxText";
            textBoxText.Size = new System.Drawing.Size(416, 20);
            textBoxText.TabIndex = 1;
            textBoxText.TextChanged += new System.EventHandler(textBoxText_TextChanged);
            textBoxText.Validating += new System.ComponentModel.CancelEventHandler(textBoxText_Validating);
            // 
            // labelPrompt
            // 
            labelPrompt.AutoSize = true;
            labelPrompt.Location = new System.Drawing.Point(15, 15);
            labelPrompt.Name = "labelPrompt";
            labelPrompt.Size = new System.Drawing.Size(39, 13);
            labelPrompt.TabIndex = 0;
            labelPrompt.Text = "prompt";
            // 
            // errorProviderText
            // 
            errorProviderText.ContainerControl = this;
            errorProviderText.DataMember = "";
            // 
            // InputBox
            // 
            AcceptButton = buttonOK;
            AutoScaleBaseSize = new System.Drawing.Size(5, 13);
            CancelButton = buttonCancel;
            ClientSize = new System.Drawing.Size(464, 104);
            Controls.Add(labelPrompt);
            Controls.Add(textBoxText);
            Controls.Add(buttonCancel);
            Controls.Add(buttonOK);
            FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            Name = "InputBox";
            RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            RightToLeftLayout = true;
            Text = "Title";
            ((System.ComponentModel.ISupportInitialize)(errorProviderText)).EndInit();
            ResumeLayout(false);
            PerformLayout();

        }
        #endregion

        private void buttonCancel_Click(object sender, System.EventArgs e)
        {
            Validator = null;
            Close();
        }

        private void buttonOK_Click(object sender, System.EventArgs e)
        {
            Close();
        }

        /// <summary>
        /// Displays a prompt in a dialog box, waits for the user to input text or click a button.
        /// </summary>
        /// <param name="prompt">String expression displayed as the message in the dialog box</param>
        /// <param name="title">String expression displayed in the title bar of the dialog box</param>
        /// <param name="defaultResponse">String expression displayed in the text box as the default response</param>
        /// <param name="validator">Delegate used to validate the text</param>
        /// <param name="xpos">Numeric expression that specifies the distance of the left edge of the dialog box from the left edge of the screen.</param>
        /// <param name="ypos">Numeric expression that specifies the distance of the upper edge of the dialog box from the top of the screen</param>
        /// <returns>An InputBoxResult object with the Text and the OK property set to true when OK was clicked.</returns>
        public static InputBoxResult Show(string prompt, string title, string defaultResponse, InputBoxValidatingHandler validator, int xpos, int ypos)
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

                InputBoxResult retval = new InputBoxResult();
                if (result == DialogResult.OK)
                {
                    retval.Text = form.textBoxText.Text;
                    retval.OK = true;
                }
                return retval;
            }
        }

        /// <summary>
        /// Displays a prompt in a dialog box, waits for the user to input text or click a button.
        /// </summary>
        /// <param name="prompt">String expression displayed as the message in the dialog box</param>
        /// <param name="title">String expression displayed in the title bar of the dialog box</param>
        /// <param name="defaultResponse">String expression displayed in the text box as the default response</param>
        /// <param name="validator">Delegate used to validate the text</param>
        /// <returns>An InputBoxResult object with the Text and the OK property set to true when OK was clicked.</returns>
        public static InputBoxResult Show(string prompt, string title, string defaultText, InputBoxValidatingHandler validator)
        {
            return Show(prompt, title, defaultText, validator, -1, -1);
        }


        /// <summary>
        /// Reset the ErrorProvider
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void textBoxText_TextChanged(object sender, System.EventArgs e)
        {
            errorProviderText.SetError(textBoxText, "");
        }

        /// <summary>
        /// Validate the Text using the Validator
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void textBoxText_Validating(object sender, System.ComponentModel.CancelEventArgs e)
        {
            if (Validator != null)
            {
                InputBoxValidatingArgs args = new InputBoxValidatingArgs();
                args.Text = textBoxText.Text;
                Validator(this, args);
                if (args.Cancel)
                {
                    e.Cancel = true;
                    errorProviderText.SetError(textBoxText, args.Message);
                }
            }
        }

        protected InputBoxValidatingHandler Validator
        {
            get
            {
                return (_validator);
            }
            set
            {
                _validator = value;
            }
        }
    }

    /// <summary>
    /// Class used to store the result of an InputBox.Show message.
    /// </summary>
    public class InputBoxResult
    {
        public bool OK;
        public string Text;
    }

    /// <summary>
    /// EventArgs used to Validate an InputBox
    /// </summary>
    public class InputBoxValidatingArgs : EventArgs
    {
        public string Text;
        public string Message;
        public bool Cancel;
    }

    /// <summary>
    /// Delegate used to Validate an InputBox
    /// </summary>
    public delegate void InputBoxValidatingHandler(object sender, InputBoxValidatingArgs e);

}
