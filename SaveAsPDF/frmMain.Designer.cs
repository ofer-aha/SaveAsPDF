
namespace SaveAsPDF
{
    partial class frmMain
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmMain));
            this.btnCancel = new System.Windows.Forms.Button();
            this.txtProjectID = new System.Windows.Forms.TextBox();
            this.lblProjectID = new System.Windows.Forms.Label();
            this.treeView1 = new System.Windows.Forms.TreeView();
            this.imageList = new System.Windows.Forms.ImageList(this.components);
            this.btnFolders = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.lblFolder = new System.Windows.Forms.Label();
            this.btnOK = new System.Windows.Forms.Button();
            this.btnSettings = new System.Windows.Forms.Button();
            this.stsStrip = new System.Windows.Forms.StatusStrip();
            this.toolStripStatusLabel1 = new System.Windows.Forms.ToolStripStatusLabel();
            this.dgvAttachments = new System.Windows.Forms.DataGridView();
            this.lblProjectName = new System.Windows.Forms.Label();
            this.txtProjectName = new System.Windows.Forms.TextBox();
            this.txtFullPath = new System.Windows.Forms.TextBox();
            this.lblEmployeeName = new System.Windows.Forms.Label();
            this.chkbSendNote = new System.Windows.Forms.CheckBox();
            this.rtxtNotes = new System.Windows.Forms.RichTextBox();
            this.lblNotes = new System.Windows.Forms.Label();
            this.btnBold = new System.Windows.Forms.Button();
            this.btnPhoneBook = new System.Windows.Forms.Button();
            this.lblAttachments = new System.Windows.Forms.Label();
            this.txtSubject = new System.Windows.Forms.TextBox();
            this.chkbSelectAllAttachments = new System.Windows.Forms.CheckBox();
            this.dlgFont = new System.Windows.Forms.FontDialog();
            this.dgvEmployees = new System.Windows.Forms.DataGridView();
            this.iD = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.firstName = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.lastName = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.emailAddress = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Column1 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dlgFolders = new System.Windows.Forms.FolderBrowserDialog();
            this.RemoveEmployee = new System.Windows.Forms.Button();
            this.stsStrip.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvAttachments)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dgvEmployees)).BeginInit();
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
            // 
            // lblProjectID
            // 
            resources.ApplyResources(this.lblProjectID, "lblProjectID");
            this.lblProjectID.Name = "lblProjectID";
            // 
            // treeView1
            // 
            resources.ApplyResources(this.treeView1, "treeView1");
            this.treeView1.ImageList = this.imageList;
            this.treeView1.Name = "treeView1";
            this.treeView1.BeforeExpand += new System.Windows.Forms.TreeViewCancelEventHandler(this.treeView1_BeforeExpand);
            // 
            // imageList
            // 
            this.imageList.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("imageList.ImageStream")));
            this.imageList.TransparentColor = System.Drawing.Color.Transparent;
            this.imageList.Images.SetKeyName(0, "Saki-NuoveXT-2-Folder-open.ico");
            this.imageList.Images.SetKeyName(1, "FolderClosed.png");
            this.imageList.Images.SetKeyName(2, "FolderOpen.png");
            this.imageList.Images.SetKeyName(3, "HardDisk.ico");
            this.imageList.Images.SetKeyName(4, "Desktop.png");
            this.imageList.Images.SetKeyName(5, "MyDocuments.png");
            this.imageList.Images.SetKeyName(6, "MyPictures.png");
            this.imageList.Images.SetKeyName(7, "MyVideos.png");
            this.imageList.Images.SetKeyName(8, "ProgramFiles.png");
            // 
            // btnFolders
            // 
            resources.ApplyResources(this.btnFolders, "btnFolders");
            this.btnFolders.Name = "btnFolders";
            this.btnFolders.UseVisualStyleBackColor = true;
            this.btnFolders.Click += new System.EventHandler(this.btnFolders_Click);
            // 
            // label1
            // 
            resources.ApplyResources(this.label1, "label1");
            this.label1.Name = "label1";
            // 
            // lblFolder
            // 
            resources.ApplyResources(this.lblFolder, "lblFolder");
            this.lblFolder.Name = "lblFolder";
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
            this.toolStripStatusLabel1});
            resources.ApplyResources(this.stsStrip, "stsStrip");
            this.stsStrip.Name = "stsStrip";
            // 
            // toolStripStatusLabel1
            // 
            this.toolStripStatusLabel1.Name = "toolStripStatusLabel1";
            resources.ApplyResources(this.toolStripStatusLabel1, "toolStripStatusLabel1");
            // 
            // dgvAttachments
            // 
            this.dgvAttachments.AllowUserToAddRows = false;
            this.dgvAttachments.AllowUserToDeleteRows = false;
            this.dgvAttachments.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells;
            this.dgvAttachments.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            resources.ApplyResources(this.dgvAttachments, "dgvAttachments");
            this.dgvAttachments.Name = "dgvAttachments";
            this.dgvAttachments.RowHeadersVisible = false;
            this.dgvAttachments.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dgvAttachments.CellDoubleClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgvAttachments_CellDoubleClick);
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
            // txtFullPath
            // 
            resources.ApplyResources(this.txtFullPath, "txtFullPath");
            this.txtFullPath.Name = "txtFullPath";
            this.txtFullPath.ReadOnly = true;
            this.txtFullPath.TabStop = false;
            // 
            // lblEmployeeName
            // 
            resources.ApplyResources(this.lblEmployeeName, "lblEmployeeName");
            this.lblEmployeeName.Name = "lblEmployeeName";
            // 
            // chkbSendNote
            // 
            resources.ApplyResources(this.chkbSendNote, "chkbSendNote");
            this.chkbSendNote.Name = "chkbSendNote";
            this.chkbSendNote.UseVisualStyleBackColor = true;
            this.chkbSendNote.CheckedChanged += new System.EventHandler(this.chkbSendNote_CheckedChanged);
            // 
            // rtxtNotes
            // 
            this.rtxtNotes.EnableAutoDragDrop = true;
            resources.ApplyResources(this.rtxtNotes, "rtxtNotes");
            this.rtxtNotes.Name = "rtxtNotes";
            // 
            // lblNotes
            // 
            resources.ApplyResources(this.lblNotes, "lblNotes");
            this.lblNotes.Name = "lblNotes";
            // 
            // btnBold
            // 
            resources.ApplyResources(this.btnBold, "btnBold");
            this.btnBold.Name = "btnBold";
            this.btnBold.UseVisualStyleBackColor = true;
            this.btnBold.Click += new System.EventHandler(this.btnStyle_Click);
            // 
            // btnPhoneBook
            // 
            resources.ApplyResources(this.btnPhoneBook, "btnPhoneBook");
            this.btnPhoneBook.Name = "btnPhoneBook";
            this.btnPhoneBook.UseVisualStyleBackColor = true;
            this.btnPhoneBook.Click += new System.EventHandler(this.btnPhoneBook_Click);
            // 
            // lblAttachments
            // 
            resources.ApplyResources(this.lblAttachments, "lblAttachments");
            this.lblAttachments.Name = "lblAttachments";
            // 
            // txtSubject
            // 
            resources.ApplyResources(this.txtSubject, "txtSubject");
            this.txtSubject.Name = "txtSubject";
            // 
            // chkbSelectAllAttachments
            // 
            resources.ApplyResources(this.chkbSelectAllAttachments, "chkbSelectAllAttachments");
            this.chkbSelectAllAttachments.Name = "chkbSelectAllAttachments";
            this.chkbSelectAllAttachments.UseVisualStyleBackColor = true;
            this.chkbSelectAllAttachments.CheckedChanged += new System.EventHandler(this.chkbSelectAllAttachments_CheckedChanged);
            // 
            // dgvEmployees
            // 
            this.dgvEmployees.AllowUserToAddRows = false;
            this.dgvEmployees.AllowUserToDeleteRows = false;
            this.dgvEmployees.AllowUserToOrderColumns = true;
            this.dgvEmployees.AllowUserToResizeRows = false;
            this.dgvEmployees.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells;
            this.dgvEmployees.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells;
            this.dgvEmployees.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvEmployees.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.iD,
            this.firstName,
            this.lastName,
            this.emailAddress,
            this.Column1});
            resources.ApplyResources(this.dgvEmployees, "dgvEmployees");
            this.dgvEmployees.Name = "dgvEmployees";
            this.dgvEmployees.RowHeadersVisible = false;
            this.dgvEmployees.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            // 
            // iD
            // 
            resources.ApplyResources(this.iD, "iD");
            this.iD.Name = "iD";
            // 
            // firstName
            // 
            resources.ApplyResources(this.firstName, "firstName");
            this.firstName.Name = "firstName";
            // 
            // lastName
            // 
            resources.ApplyResources(this.lastName, "lastName");
            this.lastName.Name = "lastName";
            // 
            // emailAddress
            // 
            resources.ApplyResources(this.emailAddress, "emailAddress");
            this.emailAddress.Name = "emailAddress";
            // 
            // Column1
            // 
            resources.ApplyResources(this.Column1, "Column1");
            this.Column1.Name = "Column1";
            // 
            // RemoveEmployee
            // 
            resources.ApplyResources(this.RemoveEmployee, "RemoveEmployee");
            this.RemoveEmployee.Name = "RemoveEmployee";
            this.RemoveEmployee.UseVisualStyleBackColor = true;
            this.RemoveEmployee.Click += new System.EventHandler(this.RemoveEmployee_Click);
            // 
            // frmMain
            // 
            resources.ApplyResources(this, "$this");
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.btnCancel;
            this.Controls.Add(this.RemoveEmployee);
            this.Controls.Add(this.dgvEmployees);
            this.Controls.Add(this.chkbSelectAllAttachments);
            this.Controls.Add(this.txtSubject);
            this.Controls.Add(this.lblAttachments);
            this.Controls.Add(this.btnPhoneBook);
            this.Controls.Add(this.btnBold);
            this.Controls.Add(this.lblNotes);
            this.Controls.Add(this.rtxtNotes);
            this.Controls.Add(this.chkbSendNote);
            this.Controls.Add(this.lblEmployeeName);
            this.Controls.Add(this.txtFullPath);
            this.Controls.Add(this.lblProjectName);
            this.Controls.Add(this.txtProjectName);
            this.Controls.Add(this.dgvAttachments);
            this.Controls.Add(this.stsStrip);
            this.Controls.Add(this.btnSettings);
            this.Controls.Add(this.btnOK);
            this.Controls.Add(this.lblFolder);
            this.Controls.Add(this.btnFolders);
            this.Controls.Add(this.treeView1);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.lblProjectID);
            this.Controls.Add(this.txtProjectID);
            this.Controls.Add(this.btnCancel);
            this.HelpButton = true;
            this.MaximizeBox = false;
            this.Name = "frmMain";
            this.Load += new System.EventHandler(this.frmMain_Load);
            this.stsStrip.ResumeLayout(false);
            this.stsStrip.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvAttachments)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dgvEmployees)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.TextBox txtProjectID;
        private System.Windows.Forms.Label lblProjectID;
        private System.Windows.Forms.TreeView treeView1;
        private System.Windows.Forms.Button btnFolders;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label lblFolder;
        private System.Windows.Forms.Button btnOK;
        private System.Windows.Forms.Button btnSettings;
        private System.Windows.Forms.StatusStrip stsStrip;
        private System.Windows.Forms.ToolStripStatusLabel toolStripStatusLabel1;
        private System.Windows.Forms.DataGridView dgvAttachments;
        private System.Windows.Forms.Label lblProjectName;
        private System.Windows.Forms.TextBox txtProjectName;
        private System.Windows.Forms.TextBox txtFullPath;
        private System.Windows.Forms.Label lblEmployeeName;
        private System.Windows.Forms.CheckBox chkbSendNote;
        private System.Windows.Forms.RichTextBox rtxtNotes;
        private System.Windows.Forms.Label lblNotes;
        private System.Windows.Forms.Button btnBold;
        private System.Windows.Forms.Button btnPhoneBook;
        private System.Windows.Forms.Label lblAttachments;
        private System.Windows.Forms.TextBox txtSubject;
        private System.Windows.Forms.CheckBox chkbSelectAllAttachments;
        private System.Windows.Forms.FontDialog dlgFont;
        private System.Windows.Forms.DataGridView dgvEmployees;
        private System.Windows.Forms.DataGridViewTextBoxColumn iD;
        private System.Windows.Forms.DataGridViewTextBoxColumn firstName;
        private System.Windows.Forms.DataGridViewTextBoxColumn lastName;
        private System.Windows.Forms.DataGridViewTextBoxColumn emailAddress;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column1;
        private System.Windows.Forms.FolderBrowserDialog dlgFolders;
        private System.Windows.Forms.ImageList imageList;
        private System.Windows.Forms.Button RemoveEmployee;
    }
}