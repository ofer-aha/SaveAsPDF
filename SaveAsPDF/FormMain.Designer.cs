namespace SaveAsPDF
{
    partial class FormMain
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
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
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FormMain));
            this.btnCancel = new System.Windows.Forms.Button();
            this.txtProjectID = new System.Windows.Forms.TextBox();
            this.lblProjectID = new System.Windows.Forms.Label();
            this.imageList = new System.Windows.Forms.ImageList(this.components);
            this.lblSubject = new System.Windows.Forms.Label();
            this.btnOK = new System.Windows.Forms.Button();
            this.btnSettings = new System.Windows.Forms.Button();
            this.stsStrip = new System.Windows.Forms.StatusStrip();
            this.tsslStatus = new System.Windows.Forms.ToolStripStatusLabel();
            this.lblProjectName = new System.Windows.Forms.Label();
            this.txtProjectName = new System.Windows.Forms.TextBox();
            this.txtSubject = new System.Windows.Forms.TextBox();
            this.dlgFont = new System.Windows.Forms.FontDialog();
            this.dlgFolders = new System.Windows.Forms.FolderBrowserDialog();
            this.btnNewProject = new System.Windows.Forms.Button();
            this.tabNotes = new System.Windows.Forms.TabControl();
            this.tabProjectNote = new System.Windows.Forms.TabPage();
            this.rtxtProjectNotes = new System.Windows.Forms.RichTextBox();
            this.btnCopyNotesToMail = new System.Windows.Forms.Button();
            this.tabMailNotes = new System.Windows.Forms.TabPage();
            this.btnCopyNotesToProject = new System.Windows.Forms.Button();
            this.btnStyle = new System.Windows.Forms.Button();
            this.rtxtNotes = new System.Windows.Forms.RichTextBox();
            this.groupBoxEmployee = new System.Windows.Forms.GroupBox();
            this.RemoveEmployee = new System.Windows.Forms.Button();
            this.dgvEmployees = new System.Windows.Forms.DataGridView();
            this.btnPhoneBook = new System.Windows.Forms.Button();
            this.chkbSendNote = new System.Windows.Forms.CheckBox();
            this.tabFilesFolders = new System.Windows.Forms.TabControl();
            this.tabFolsers = new System.Windows.Forms.TabPage();
            this.label1 = new System.Windows.Forms.Label();
            this.txtFullPath = new System.Windows.Forms.TextBox();
            this.btnFolders = new System.Windows.Forms.Button();
            this.tvFolders = new System.Windows.Forms.TreeView();
            this.tabAtachments = new System.Windows.Forms.TabPage();
            this.chkbSelectAllAttachments = new System.Windows.Forms.CheckBox();
            this.dgvAttachments = new System.Windows.Forms.DataGridView();
            this.lblSaveLocation = new System.Windows.Forms.Label();
            this.chbOpenPDF = new System.Windows.Forms.CheckBox();
            this.errorProviderMain = new System.Windows.Forms.ErrorProvider(this.components);
            this.cmbSaveLocation = new System.Windows.Forms.ComboBox();
            this.stsStrip.SuspendLayout();
            this.tabNotes.SuspendLayout();
            this.tabProjectNote.SuspendLayout();
            this.tabMailNotes.SuspendLayout();
            this.groupBoxEmployee.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvEmployees)).BeginInit();
            this.tabFilesFolders.SuspendLayout();
            this.tabFolsers.SuspendLayout();
            this.tabAtachments.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvAttachments)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.errorProviderMain)).BeginInit();
            this.SuspendLayout();
            // 
            // btnCancel
            // 
            resources.ApplyResources(this.btnCancel, "btnCancel");
            this.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.UseVisualStyleBackColor = true;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // txtProjectID
            // 
            resources.ApplyResources(this.txtProjectID, "txtProjectID");
            this.txtProjectID.Name = "txtProjectID";
            this.txtProjectID.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtProjectID_KeyDown);
            this.txtProjectID.MouseHover += new System.EventHandler(this.txtProjectID_MouseHover);
            this.txtProjectID.Validating += new System.ComponentModel.CancelEventHandler(this.txtProjectID_Validating);
            this.txtProjectID.Validated += new System.EventHandler(this.txtProjectID_Validated);
            // 
            // lblProjectID
            // 
            resources.ApplyResources(this.lblProjectID, "lblProjectID");
            this.lblProjectID.Name = "lblProjectID";
            // 
            // imageList
            // 
            this.imageList.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("imageList.ImageStream")));
            this.imageList.TransparentColor = System.Drawing.Color.Transparent;
            this.imageList.Images.SetKeyName(0, "FolderOpen.png");
            this.imageList.Images.SetKeyName(1, "FolderClosed.png");
            this.imageList.Images.SetKeyName(2, "HardDisk.ico");
            this.imageList.Images.SetKeyName(3, "Desktop.png");
            this.imageList.Images.SetKeyName(4, "MyDocuments.png");
            this.imageList.Images.SetKeyName(5, "MyPictures.png");
            this.imageList.Images.SetKeyName(6, "MyVideos.png");
            this.imageList.Images.SetKeyName(7, "ProgramFiles.png");
            // 
            // lblSubject
            // 
            resources.ApplyResources(this.lblSubject, "lblSubject");
            this.lblSubject.Name = "lblSubject";
            // 
            // btnOK
            // 
            resources.ApplyResources(this.btnOK, "btnOK");
            this.btnOK.Name = "btnOK";
            this.btnOK.UseVisualStyleBackColor = true;
            this.btnOK.Click += new System.EventHandler(this.btnOK_Click);
            // 
            // btnSettings
            // 
            resources.ApplyResources(this.btnSettings, "btnSettings");
            this.btnSettings.Name = "btnSettings";
            this.btnSettings.UseVisualStyleBackColor = true;
            this.btnSettings.Click += new System.EventHandler(this.BtnSettings_Click);
            // 
            // stsStrip
            // 
            this.stsStrip.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.tsslStatus});
            resources.ApplyResources(this.stsStrip, "stsStrip");
            this.stsStrip.Name = "stsStrip";
            // 
            // tsslStatus
            // 
            this.tsslStatus.Name = "tsslStatus";
            resources.ApplyResources(this.tsslStatus, "tsslStatus");
            // 
            // lblProjectName
            // 
            resources.ApplyResources(this.lblProjectName, "lblProjectName");
            this.lblProjectName.Name = "lblProjectName";
            // 
            // txtProjectName
            // 
            resources.ApplyResources(this.txtProjectName, "txtProjectName");
            this.txtProjectName.Name = "txtProjectName";
            // 
            // txtSubject
            // 
            resources.ApplyResources(this.txtSubject, "txtSubject");
            this.txtSubject.Name = "txtSubject";
            this.txtSubject.ReadOnly = true;
            this.txtSubject.TabStop = false;
            // 
            // btnNewProject
            // 
            resources.ApplyResources(this.btnNewProject, "btnNewProject");
            this.btnNewProject.Name = "btnNewProject";
            this.btnNewProject.UseVisualStyleBackColor = true;
            this.btnNewProject.Click += new System.EventHandler(this.btnNewProject_Click);
            // 
            // tabNotes
            // 
            this.tabNotes.Controls.Add(this.tabProjectNote);
            this.tabNotes.Controls.Add(this.tabMailNotes);
            resources.ApplyResources(this.tabNotes, "tabNotes");
            this.tabNotes.Name = "tabNotes";
            this.tabNotes.SelectedIndex = 0;
            // 
            // tabProjectNote
            // 
            this.tabProjectNote.Controls.Add(this.rtxtProjectNotes);
            this.tabProjectNote.Controls.Add(this.btnCopyNotesToMail);
            resources.ApplyResources(this.tabProjectNote, "tabProjectNote");
            this.tabProjectNote.Name = "tabProjectNote";
            this.tabProjectNote.UseVisualStyleBackColor = true;
            // 
            // rtxtProjectNotes
            // 
            resources.ApplyResources(this.rtxtProjectNotes, "rtxProjectNotes");
            this.rtxtProjectNotes.Name = "rtxtProjectNotes";
            // 
            // btnCopyNotesToMail
            // 
            resources.ApplyResources(this.btnCopyNotesToMail, "btnCopyNotesToMail");
            this.btnCopyNotesToMail.Name = "btnCopyNotesToMail";
            this.btnCopyNotesToMail.UseVisualStyleBackColor = true;
            this.btnCopyNotesToMail.Click += new System.EventHandler(this.btnCopyNotesToMail_Click);
            // 
            // tabMailNotes
            // 
            this.tabMailNotes.Controls.Add(this.btnCopyNotesToProject);
            this.tabMailNotes.Controls.Add(this.btnStyle);
            this.tabMailNotes.Controls.Add(this.rtxtNotes);
            resources.ApplyResources(this.tabMailNotes, "tabMailNotes");
            this.tabMailNotes.Name = "tabMailNotes";
            this.tabMailNotes.UseVisualStyleBackColor = true;
            // 
            // btnCopyNotesToProject
            // 
            resources.ApplyResources(this.btnCopyNotesToProject, "btnCopyNotesToProject");
            this.btnCopyNotesToProject.Name = "btnCopyNotesToProject";
            this.btnCopyNotesToProject.UseVisualStyleBackColor = true;
            this.btnCopyNotesToProject.Click += new System.EventHandler(this.btnCopyNotesToProject_Click);
            // 
            // btnStyle
            // 
            resources.ApplyResources(this.btnStyle, "btnStyle");
            this.btnStyle.Name = "btnStyle";
            this.btnStyle.UseVisualStyleBackColor = true;
            this.btnStyle.Click += new System.EventHandler(this.btnStyle_Click);
            // 
            // rtxtNotes
            // 
            this.rtxtNotes.EnableAutoDragDrop = true;
            resources.ApplyResources(this.rtxtNotes, "rtxNotes");
            this.rtxtNotes.Name = "rtxtNotes";
            // 
            // groupBoxEmployee
            // 
            this.groupBoxEmployee.Controls.Add(this.RemoveEmployee);
            this.groupBoxEmployee.Controls.Add(this.dgvEmployees);
            this.groupBoxEmployee.Controls.Add(this.btnPhoneBook);
            this.groupBoxEmployee.Controls.Add(this.chkbSendNote);
            resources.ApplyResources(this.groupBoxEmployee, "groupBoxEmployee");
            this.groupBoxEmployee.Name = "groupBoxEmployee";
            this.groupBoxEmployee.TabStop = false;
            // 
            // RemoveEmployee
            // 
            resources.ApplyResources(this.RemoveEmployee, "RemoveEmployee");
            this.RemoveEmployee.Name = "RemoveEmployee";
            this.RemoveEmployee.UseVisualStyleBackColor = true;
            this.RemoveEmployee.Click += new System.EventHandler(this.RemoveEmployee_Click);
            // 
            // dgvEmployees
            // 
            this.dgvEmployees.AllowUserToAddRows = false;
            this.dgvEmployees.AllowUserToDeleteRows = false;
            this.dgvEmployees.AllowUserToResizeRows = false;
            this.dgvEmployees.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells;
            this.dgvEmployees.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells;
            this.dgvEmployees.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            resources.ApplyResources(this.dgvEmployees, "dgvEmployees");
            this.dgvEmployees.Name = "dgvEmployees";
            this.dgvEmployees.ReadOnly = true;
            this.dgvEmployees.RowHeadersVisible = false;
            this.dgvEmployees.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            // 
            // btnPhoneBook
            // 
            resources.ApplyResources(this.btnPhoneBook, "btnPhoneBook");
            this.btnPhoneBook.Name = "btnPhoneBook";
            this.btnPhoneBook.UseVisualStyleBackColor = true;
            this.btnPhoneBook.Click += new System.EventHandler(this.btnPhoneBook_Click);
            // 
            // chkbSendNote
            // 
            resources.ApplyResources(this.chkbSendNote, "chkbSendNote");
            this.chkbSendNote.Name = "chkbSendNote";
            this.chkbSendNote.UseVisualStyleBackColor = true;
            // 
            // tabFilesFolders
            // 
            this.tabFilesFolders.Controls.Add(this.tabFolsers);
            this.tabFilesFolders.Controls.Add(this.tabAtachments);
            resources.ApplyResources(this.tabFilesFolders, "tabFilesFolders");
            this.tabFilesFolders.Name = "tabFilesFolders";
            this.tabFilesFolders.SelectedIndex = 0;
            // 
            // tabFolsers
            // 
            this.tabFolsers.Controls.Add(this.label1);
            this.tabFolsers.Controls.Add(this.txtFullPath);
            this.tabFolsers.Controls.Add(this.btnFolders);
            this.tabFolsers.Controls.Add(this.tvFolders);
            resources.ApplyResources(this.tabFolsers, "tabFolsers");
            this.tabFolsers.Name = "tabFolsers";
            this.tabFolsers.UseVisualStyleBackColor = true;
            // 
            // label1
            // 
            resources.ApplyResources(this.label1, "label1");
            this.label1.Name = "label1";
            // 
            // txtFullPath
            // 
            resources.ApplyResources(this.txtFullPath, "txtFullPath");
            this.txtFullPath.Name = "txtFullPath";
            this.txtFullPath.ReadOnly = true;
            this.txtFullPath.TabStop = false;
            // 
            // btnFolders
            // 
            resources.ApplyResources(this.btnFolders, "btnFolders");
            this.btnFolders.Name = "btnFolders";
            this.btnFolders.UseVisualStyleBackColor = true;
            this.btnFolders.Click += new System.EventHandler(this.btnFolders_Click);
            // 
            // tvFolders
            // 
            resources.ApplyResources(this.tvFolders, "tvFolders");
            this.tvFolders.ImageList = this.imageList;
            this.tvFolders.Name = "tvFolders";
            this.tvFolders.LabelEdit = true;
            this.tvFolders.AfterLabelEdit += new System.Windows.Forms.NodeLabelEditEventHandler(this.tvFolders_AfterLabelEdit);
            this.tvFolders.BeforeCollapse += new System.Windows.Forms.TreeViewCancelEventHandler(this.tvFolders_BeforeCollapse);
            this.tvFolders.BeforeExpand += new System.Windows.Forms.TreeViewCancelEventHandler(this.tvFolders_BeforeExpand);
            this.tvFolders.AfterSelect += new System.Windows.Forms.TreeViewEventHandler(this.tvFolders_AfterSelect);
            this.tvFolders.NodeMouseClick += new System.Windows.Forms.TreeNodeMouseClickEventHandler(this.tvFolders_NodeMouseClick);
            this.tvFolders.NodeMouseDoubleClick += new System.Windows.Forms.TreeNodeMouseClickEventHandler(this.tvFolders_NodeMouseDoubleClick);
            this.tvFolders.MouseDown += new System.Windows.Forms.MouseEventHandler(this.tvFolders_MouseDown);
            // 
            // tabAtachments
            // 
            this.tabAtachments.Controls.Add(this.chkbSelectAllAttachments);
            this.tabAtachments.Controls.Add(this.dgvAttachments);
            this.tabAtachments.Controls.Add(this.lblSaveLocation);
            this.tabAtachments.Controls.Add(this.chbOpenPDF);
            resources.ApplyResources(this.tabAtachments, "tabAtachments");
            this.tabAtachments.Name = "tabAtachments";
            this.tabAtachments.UseVisualStyleBackColor = true;
            // 
            // chkbSelectAllAttachments
            // 
            resources.ApplyResources(this.chkbSelectAllAttachments, "chkbSelectAllAttachments");
            this.chkbSelectAllAttachments.Name = "chkbSelectAllAttachments";
            this.chkbSelectAllAttachments.UseVisualStyleBackColor = true;
            this.chkbSelectAllAttachments.CheckedChanged += new System.EventHandler(this.chkbSelectAllAttachments_CheckedChanged);
            // 
            // dgvAttachments
            // 
            this.dgvAttachments.AllowUserToAddRows = false;
            this.dgvAttachments.AllowUserToDeleteRows = false;
            this.dgvAttachments.AllowUserToResizeColumns = false;
            this.dgvAttachments.AllowUserToResizeRows = false;
            this.dgvAttachments.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells;
            this.dgvAttachments.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            resources.ApplyResources(this.dgvAttachments, "dgvAttachments");
            this.dgvAttachments.Name = "dgvAttachments";
            this.dgvAttachments.RowHeadersVisible = false;
            this.dgvAttachments.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dgvAttachments.CellDoubleClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgvAttachments_CellDoubleClick);
            // 
            // lblSaveLocation
            // 
            resources.ApplyResources(this.lblSaveLocation, "lblSaveLocation");
            this.lblSaveLocation.Name = "lblSaveLocation";
            // 
            // chbOpenPDF
            // 
            resources.ApplyResources(this.chbOpenPDF, "chbOpenPDF");
            this.chbOpenPDF.Name = "chbOpenPDF";
            this.chbOpenPDF.UseVisualStyleBackColor = true;
            this.chbOpenPDF.CheckedChanged += new System.EventHandler(this.chbOpenPDF_CheckedChanged);
            // 
            // errorProviderMain
            // 
            this.errorProviderMain.ContainerControl = this;
            resources.ApplyResources(this.errorProviderMain, "errorProviderMain");
            // 
            // cmbSaveLocation
            // 
            this.cmbSaveLocation.FormattingEnabled = true;
            resources.ApplyResources(this.cmbSaveLocation, "cmbSaveLocation");
            this.cmbSaveLocation.Name = "cmbSaveLocation";
            this.cmbSaveLocation.TextUpdate += new System.EventHandler(this.cmbSaveLocation_TextUpdate);
            this.cmbSaveLocation.SelectedValueChanged += new System.EventHandler(this.cmbSaveLocation_SelectedValueChanged);
            // 
            // FormMain
            // 
            resources.ApplyResources(this, "$this");
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.btnCancel;
            this.Controls.Add(this.cmbSaveLocation);
            this.Controls.Add(this.chbOpenPDF);
            this.Controls.Add(this.lblSaveLocation);
            this.Controls.Add(this.tabFilesFolders);
            this.Controls.Add(this.groupBoxEmployee);
            this.Controls.Add(this.tabNotes);
            this.Controls.Add(this.btnNewProject);
            this.Controls.Add(this.txtSubject);
            this.Controls.Add(this.lblProjectName);
            this.Controls.Add(this.txtProjectName);
            this.Controls.Add(this.stsStrip);
            this.Controls.Add(this.btnSettings);
            this.Controls.Add(this.btnOK);
            this.Controls.Add(this.lblSubject);
            this.Controls.Add(this.lblProjectID);
            this.Controls.Add(this.txtProjectID);
            this.Controls.Add(this.btnCancel);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D;
            this.HelpButton = true;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "FormMain";
            this.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide;
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.FormMain_FormClosing);
            this.Load += new System.EventHandler(this.FormMain_Load);
            this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.FormMain_KeyDown);
            this.stsStrip.ResumeLayout(false);
            this.stsStrip.PerformLayout();
            this.tabNotes.ResumeLayout(false);
            this.tabProjectNote.ResumeLayout(false);
            this.tabMailNotes.ResumeLayout(false);
            this.groupBoxEmployee.ResumeLayout(false);
            this.groupBoxEmployee.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvEmployees)).EndInit();
            this.tabFilesFolders.ResumeLayout(false);
            this.tabFolsers.ResumeLayout(false);
            this.tabFolsers.PerformLayout();
            this.tabAtachments.ResumeLayout(false);
            this.tabAtachments.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvAttachments)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.errorProviderMain)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.TextBox txtProjectID;
        private System.Windows.Forms.Label lblProjectID;
        private System.Windows.Forms.Label lblSubject;
        private System.Windows.Forms.Button btnOK;
        private System.Windows.Forms.Button btnSettings;
        private System.Windows.Forms.StatusStrip stsStrip;
        private System.Windows.Forms.Label lblProjectName;
        private System.Windows.Forms.TextBox txtProjectName;
        private System.Windows.Forms.TextBox txtSubject;
        private System.Windows.Forms.FontDialog dlgFont;
        private System.Windows.Forms.FolderBrowserDialog dlgFolders;
        private System.Windows.Forms.ImageList imageList;
        private System.Windows.Forms.Button btnNewProject;
        private System.Windows.Forms.ToolStripStatusLabel tsslStatus;
        private System.Windows.Forms.TabControl tabNotes;
        private System.Windows.Forms.TabPage tabProjectNote;
        private System.Windows.Forms.TabPage tabMailNotes;
        private System.Windows.Forms.Button btnCopyNotesToMail;
        private System.Windows.Forms.Button btnCopyNotesToProject;
        private System.Windows.Forms.Button btnStyle;
        private System.Windows.Forms.RichTextBox rtxtNotes;
        private System.Windows.Forms.RichTextBox rtxtProjectNotes;
        private System.Windows.Forms.GroupBox groupBoxEmployee;
        private System.Windows.Forms.Button RemoveEmployee;
        private System.Windows.Forms.DataGridView dgvEmployees;
        private System.Windows.Forms.Button btnPhoneBook;
        private System.Windows.Forms.CheckBox chkbSendNote;
        private System.Windows.Forms.TabControl tabFilesFolders;
        private System.Windows.Forms.TabPage tabFolsers;
        private System.Windows.Forms.TextBox txtFullPath;
        private System.Windows.Forms.Button btnFolders;
        private System.Windows.Forms.TreeView tvFolders;
        private System.Windows.Forms.TabPage tabAtachments;
        private System.Windows.Forms.CheckBox chkbSelectAllAttachments;
        private System.Windows.Forms.DataGridView dgvAttachments;
        private System.Windows.Forms.Label lblSaveLocation;
        private System.Windows.Forms.CheckBox chbOpenPDF;
        private System.Windows.Forms.ErrorProvider errorProviderMain;
        private System.Windows.Forms.ComboBox cmbSaveLocation;
        private System.Windows.Forms.Label label1;
    }
}