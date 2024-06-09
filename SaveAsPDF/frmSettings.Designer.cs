
namespace SaveAsPDF
{
    partial class frmSettings
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmSettings));
            this.bntCancel = new System.Windows.Forms.Button();
            this.menuTree = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.menueAdd = new System.Windows.Forms.ToolStripMenuItem();
            this.menuAddDate = new System.Windows.Forms.ToolStripMenuItem();
            this.menuDel = new System.Windows.Forms.ToolStripMenuItem();
            this.menuRename = new System.Windows.Forms.ToolStripMenuItem();
            this.menuAppendDate = new System.Windows.Forms.ToolStripMenuItem();
            this.imageList = new System.Windows.Forms.ImageList(this.components);
            this.btnOK = new System.Windows.Forms.Button();
            this.txtRootFolder = new System.Windows.Forms.TextBox();
            this.lblRootFolder = new System.Windows.Forms.Label();
            this.btnSaveSettings = new System.Windows.Forms.Button();
            this.dlgFolders = new System.Windows.Forms.FolderBrowserDialog();
            this.btnFolders = new System.Windows.Forms.Button();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.btnLoadDefaultTree = new System.Windows.Forms.Button();
            this.tvProjectSubFolders = new System.Windows.Forms.TreeView();
            this.btnSaveTreeFile = new System.Windows.Forms.Button();
            this.btnLoadTreeFile = new System.Windows.Forms.Button();
            this.btnSaveDefaultTree = new System.Windows.Forms.Button();
            this.groupBoxDefaultFolder = new System.Windows.Forms.GroupBox();
            this.cmbDefaultFolder = new System.Windows.Forms.ComboBox();
            this.gbAttaments = new System.Windows.Forms.GroupBox();
            this.textBox2 = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.txtMinAttSize = new System.Windows.Forms.TextBox();
            this.lblMinAttSize = new System.Windows.Forms.Label();
            this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
            this.menuTree.SuspendLayout();
            this.groupBox1.SuspendLayout();
            this.groupBoxDefaultFolder.SuspendLayout();
            this.gbAttaments.SuspendLayout();
            this.SuspendLayout();
            // 
            // bntCancel
            // 
            this.bntCancel.Anchor = System.Windows.Forms.AnchorStyles.Bottom;
            this.bntCancel.Location = new System.Drawing.Point(487, 296);
            this.bntCancel.Name = "bntCancel";
            this.bntCancel.Size = new System.Drawing.Size(75, 23);
            this.bntCancel.TabIndex = 2;
            this.bntCancel.Text = "ביטול";
            this.bntCancel.UseVisualStyleBackColor = true;
            this.bntCancel.Click += new System.EventHandler(this.bntCancel_Click);
            // 
            // menuTree
            // 
            this.menuTree.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.menueAdd,
            this.menuAddDate,
            this.menuDel,
            this.menuRename,
            this.menuAppendDate});
            this.menuTree.Name = "cMenu_Add";
            this.menuTree.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.menuTree.Size = new System.Drawing.Size(202, 114);
            // 
            // menueAdd
            // 
            this.menueAdd.Image = global::SaveAsPDF.Properties.Resources.FolderClose;
            this.menueAdd.Name = "menueAdd";
            this.menueAdd.Size = new System.Drawing.Size(201, 22);
            this.menueAdd.Text = "הוסף תיקייה";
            this.menueAdd.Click += new System.EventHandler(this.menueAdd_Click);
            // 
            // menuAddDate
            // 
            this.menuAddDate.Name = "menuAddDate";
            this.menuAddDate.Size = new System.Drawing.Size(201, 22);
            this.menuAddDate.Text = "הוסף תאריך";
            this.menuAddDate.Click += new System.EventHandler(this.menuAddDate_Click);
            // 
            // menuDel
            // 
            this.menuDel.Image = global::SaveAsPDF.Properties.Resources.close_big;
            this.menuDel.Name = "menuDel";
            this.menuDel.Size = new System.Drawing.Size(201, 22);
            this.menuDel.Text = "מחק תיקייה";
            this.menuDel.Click += new System.EventHandler(this.menuDel_Click);
            // 
            // menuRename
            // 
            this.menuRename.Name = "menuRename";
            this.menuRename.Size = new System.Drawing.Size(201, 22);
            this.menuRename.Text = "שנה שם";
            this.menuRename.Click += new System.EventHandler(this.menuRename_Click);
            // 
            // menuAppendDate
            // 
            this.menuAppendDate.Name = "menuAppendDate";
            this.menuAppendDate.Size = new System.Drawing.Size(201, 22);
            this.menuAppendDate.Text = "הוסף תאריך לשם תיקייה";
            this.menuAppendDate.Click += new System.EventHandler(this.menuAppendDate_Click);
            // 
            // imageList
            // 
            this.imageList.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("imageList.ImageStream")));
            this.imageList.TransparentColor = System.Drawing.Color.Transparent;
            this.imageList.Images.SetKeyName(0, "FolderClosed.png");
            this.imageList.Images.SetKeyName(1, "Saki-NuoveXT-2-Folder-open.ico");
            this.imageList.Images.SetKeyName(2, "FolderOpen.png");
            this.imageList.Images.SetKeyName(3, "HardDisk.ico");
            this.imageList.Images.SetKeyName(4, "Desktop.png");
            this.imageList.Images.SetKeyName(5, "MyDocuments.png");
            this.imageList.Images.SetKeyName(6, "MyPictures.png");
            this.imageList.Images.SetKeyName(7, "MyVideos.png");
            this.imageList.Images.SetKeyName(8, "ProgramFiles.png");
            // 
            // btnOK
            // 
            this.btnOK.Anchor = System.Windows.Forms.AnchorStyles.Bottom;
            this.btnOK.Location = new System.Drawing.Point(225, 296);
            this.btnOK.Name = "btnOK";
            this.btnOK.Size = new System.Drawing.Size(75, 23);
            this.btnOK.TabIndex = 0;
            this.btnOK.Text = "אישור";
            this.btnOK.UseVisualStyleBackColor = true;
            this.btnOK.Click += new System.EventHandler(this.btnOK_Click);
            // 
            // txtRootFolder
            // 
            this.txtRootFolder.Location = new System.Drawing.Point(225, 29);
            this.txtRootFolder.Name = "txtRootFolder";
            this.txtRootFolder.ReadOnly = true;
            this.txtRootFolder.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.txtRootFolder.Size = new System.Drawing.Size(121, 20);
            this.txtRootFolder.TabIndex = 4;
            this.txtRootFolder.TextChanged += new System.EventHandler(this.txtRootFolder_TextChanged);
            // 
            // lblRootFolder
            // 
            this.lblRootFolder.AutoSize = true;
            this.lblRootFolder.Location = new System.Drawing.Point(16, 33);
            this.lblRootFolder.Name = "lblRootFolder";
            this.lblRootFolder.Size = new System.Drawing.Size(134, 13);
            this.lblRootFolder.TabIndex = 5;
            this.lblRootFolder.Text = "תיקיית שורש לפרויקטים";
            // 
            // btnSaveSettings
            // 
            this.btnSaveSettings.Anchor = System.Windows.Forms.AnchorStyles.Bottom;
            this.btnSaveSettings.Location = new System.Drawing.Point(354, 296);
            this.btnSaveSettings.Name = "btnSaveSettings";
            this.btnSaveSettings.Size = new System.Drawing.Size(97, 23);
            this.btnSaveSettings.TabIndex = 1;
            this.btnSaveSettings.Text = "שמור שינויים";
            this.btnSaveSettings.UseVisualStyleBackColor = true;
            this.btnSaveSettings.Click += new System.EventHandler(this.btnSaveSettings_Click);
            // 
            // btnFolders
            // 
            this.btnFolders.Location = new System.Drawing.Point(354, 28);
            this.btnFolders.Name = "btnFolders";
            this.btnFolders.Size = new System.Drawing.Size(24, 23);
            this.btnFolders.TabIndex = 9;
            this.btnFolders.Text = "...";
            this.btnFolders.UseVisualStyleBackColor = true;
            this.btnFolders.Click += new System.EventHandler(this.btnFolders_Click);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.btnLoadDefaultTree);
            this.groupBox1.Controls.Add(this.tvProjectSubFolders);
            this.groupBox1.Controls.Add(this.btnSaveTreeFile);
            this.groupBox1.Controls.Add(this.btnLoadTreeFile);
            this.groupBox1.Controls.Add(this.btnSaveDefaultTree);
            this.groupBox1.Location = new System.Drawing.Point(427, 10);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(344, 272);
            this.groupBox1.TabIndex = 12;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "מבנה תיקיות";
            // 
            // btnLoadDefaultTree
            // 
            this.btnLoadDefaultTree.Location = new System.Drawing.Point(219, 51);
            this.btnLoadDefaultTree.Name = "btnLoadDefaultTree";
            this.btnLoadDefaultTree.Size = new System.Drawing.Size(110, 23);
            this.btnLoadDefaultTree.TabIndex = 1;
            this.btnLoadDefaultTree.Text = "טען ברירת מחדל";
            this.btnLoadDefaultTree.UseVisualStyleBackColor = true;
            this.btnLoadDefaultTree.Click += new System.EventHandler(this.btnLoadDefaultTree_Click);
            // 
            // tvProjectSubFolders
            // 
            this.tvProjectSubFolders.ContextMenuStrip = this.menuTree;
            this.tvProjectSubFolders.ImageIndex = 0;
            this.tvProjectSubFolders.ImageList = this.imageList;
            this.tvProjectSubFolders.Location = new System.Drawing.Point(6, 19);
            this.tvProjectSubFolders.Name = "tvProjectSubFolders";
            this.tvProjectSubFolders.SelectedImageIndex = 0;
            this.tvProjectSubFolders.Size = new System.Drawing.Size(207, 214);
            this.tvProjectSubFolders.TabIndex = 15;
            this.tvProjectSubFolders.AfterLabelEdit += new System.Windows.Forms.NodeLabelEditEventHandler(this.tvProjectSubFolders_AfterLabelEdit);
            this.tvProjectSubFolders.AfterSelect += new System.Windows.Forms.TreeViewEventHandler(this.tvProjectSubFolders_AfterSelect);
            this.tvProjectSubFolders.MouseDown += new System.Windows.Forms.MouseEventHandler(this.tvProjectSubFolders_MouseDown);
            // 
            // btnSaveTreeFile
            // 
            this.btnSaveTreeFile.Location = new System.Drawing.Point(219, 107);
            this.btnSaveTreeFile.Name = "btnSaveTreeFile";
            this.btnSaveTreeFile.Size = new System.Drawing.Size(110, 23);
            this.btnSaveTreeFile.TabIndex = 3;
            this.btnSaveTreeFile.Text = "יצא קובץ";
            this.btnSaveTreeFile.UseVisualStyleBackColor = true;
            this.btnSaveTreeFile.Click += new System.EventHandler(this.btnSaveTreeFile_Click);
            // 
            // btnLoadTreeFile
            // 
            this.btnLoadTreeFile.Location = new System.Drawing.Point(219, 79);
            this.btnLoadTreeFile.Name = "btnLoadTreeFile";
            this.btnLoadTreeFile.Size = new System.Drawing.Size(110, 23);
            this.btnLoadTreeFile.TabIndex = 2;
            this.btnLoadTreeFile.Text = "יבא קובץ";
            this.btnLoadTreeFile.UseVisualStyleBackColor = true;
            this.btnLoadTreeFile.Click += new System.EventHandler(this.btnLoadTreeFile_Click_1);
            // 
            // btnSaveDefaultTree
            // 
            this.btnSaveDefaultTree.Location = new System.Drawing.Point(219, 23);
            this.btnSaveDefaultTree.Name = "btnSaveDefaultTree";
            this.btnSaveDefaultTree.Size = new System.Drawing.Size(110, 23);
            this.btnSaveDefaultTree.TabIndex = 0;
            this.btnSaveDefaultTree.Text = "שמור ברירת מחדל";
            this.btnSaveDefaultTree.UseVisualStyleBackColor = true;
            this.btnSaveDefaultTree.Click += new System.EventHandler(this.btnSaveDefaultTree_Click);
            // 
            // groupBoxDefaultFolder
            // 
            this.groupBoxDefaultFolder.Controls.Add(this.cmbDefaultFolder);
            this.groupBoxDefaultFolder.Location = new System.Drawing.Point(16, 65);
            this.groupBoxDefaultFolder.Name = "groupBoxDefaultFolder";
            this.groupBoxDefaultFolder.Size = new System.Drawing.Size(362, 47);
            this.groupBoxDefaultFolder.TabIndex = 15;
            this.groupBoxDefaultFolder.TabStop = false;
            this.groupBoxDefaultFolder.Text = "תיקיית ברירת מחדל לשמירת מכתבים";
            // 
            // cmbDefaultFolder
            // 
            this.cmbDefaultFolder.FormattingEnabled = true;
            this.cmbDefaultFolder.Location = new System.Drawing.Point(8, 19);
            this.cmbDefaultFolder.Name = "cmbDefaultFolder";
            this.cmbDefaultFolder.Size = new System.Drawing.Size(345, 21);
            this.cmbDefaultFolder.TabIndex = 21;
            this.cmbDefaultFolder.SelectedIndexChanged += new System.EventHandler(this.cmbDefaultFolder_SelectedIndexChanged);
            // 
            // gbAttaments
            // 
            this.gbAttaments.Controls.Add(this.textBox2);
            this.gbAttaments.Controls.Add(this.label3);
            this.gbAttaments.Controls.Add(this.txtMinAttSize);
            this.gbAttaments.Controls.Add(this.lblMinAttSize);
            this.gbAttaments.Location = new System.Drawing.Point(26, 183);
            this.gbAttaments.Name = "gbAttaments";
            this.gbAttaments.Size = new System.Drawing.Size(350, 99);
            this.gbAttaments.TabIndex = 16;
            this.gbAttaments.TabStop = false;
            this.gbAttaments.Text = "קבצים מצורפים";
            // 
            // textBox2
            // 
            this.textBox2.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.textBox2.Location = new System.Drawing.Point(6, 46);
            this.textBox2.Multiline = true;
            this.textBox2.Name = "textBox2";
            this.textBox2.ReadOnly = true;
            this.textBox2.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.textBox2.Size = new System.Drawing.Size(338, 47);
            this.textBox2.TabIndex = 18;
            this.textBox2.Text = "קבצים קטנים בד\"כ מהווים חלק מהחתימה הקבועה במייל ורצוי להתעלם מהם\r\nbytes  8192: ע" +
    "רך מומלץ\r\n";
            this.textBox2.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(-55, 47);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(0, 13);
            this.label3.TabIndex = 3;
            // 
            // txtMinAttSize
            // 
            this.txtMinAttSize.Location = new System.Drawing.Point(6, 20);
            this.txtMinAttSize.Name = "txtMinAttSize";
            this.txtMinAttSize.Size = new System.Drawing.Size(80, 20);
            this.txtMinAttSize.TabIndex = 1;
            this.txtMinAttSize.TextChanged += new System.EventHandler(this.txtMinAttSize_TextChanged);
            // 
            // lblMinAttSize
            // 
            this.lblMinAttSize.AutoSize = true;
            this.lblMinAttSize.Location = new System.Drawing.Point(170, 27);
            this.lblMinAttSize.Name = "lblMinAttSize";
            this.lblMinAttSize.Size = new System.Drawing.Size(177, 13);
            this.lblMinAttSize.TabIndex = 0;
            this.lblMinAttSize.Text = "גודל מינימאלי לקבצים מצורפים:";
            // 
            // frmSettings
            // 
            this.ClientSize = new System.Drawing.Size(786, 327);
            this.Controls.Add(this.gbAttaments);
            this.Controls.Add(this.groupBoxDefaultFolder);
            this.Controls.Add(this.btnFolders);
            this.Controls.Add(this.btnSaveSettings);
            this.Controls.Add(this.lblRootFolder);
            this.Controls.Add(this.txtRootFolder);
            this.Controls.Add(this.btnOK);
            this.Controls.Add(this.bntCancel);
            this.Controls.Add(this.groupBox1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "frmSettings";
            this.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.RightToLeftLayout = true;
            this.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "הגדרות";
            this.TopMost = true;
            this.Load += new System.EventHandler(this.frmSettings_Load);
            this.menuTree.ResumeLayout(false);
            this.groupBox1.ResumeLayout(false);
            this.groupBoxDefaultFolder.ResumeLayout(false);
            this.gbAttaments.ResumeLayout(false);
            this.gbAttaments.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        
        private System.Windows.Forms.Button bntCancel;
        private System.Windows.Forms.ContextMenuStrip menuTree;
        private System.Windows.Forms.ToolStripMenuItem menueAdd;
        private System.Windows.Forms.ToolStripMenuItem menuDel;
        private System.Windows.Forms.ToolStripMenuItem menuRename;
        private System.Windows.Forms.Button btnOK;
        private System.Windows.Forms.ImageList imageList;
        private System.Windows.Forms.TextBox txtRootFolder;
        private System.Windows.Forms.Label lblRootFolder;
        private System.Windows.Forms.Button btnSaveSettings;
        private System.Windows.Forms.FolderBrowserDialog dlgFolders;
        private System.Windows.Forms.Button btnFolders;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.TreeView tvProjectSubFolders;
        private System.Windows.Forms.Button btnSaveTreeFile;
        private System.Windows.Forms.Button btnLoadTreeFile;
        private System.Windows.Forms.Button btnSaveDefaultTree;
        private System.Windows.Forms.Button btnLoadDefaultTree;
        private System.Windows.Forms.GroupBox groupBoxDefaultFolder;
        private System.Windows.Forms.GroupBox gbAttaments;
        private System.Windows.Forms.TextBox txtMinAttSize;
        private System.Windows.Forms.Label lblMinAttSize;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox textBox2;
        private System.Windows.Forms.ToolTip toolTip1;
        private System.Windows.Forms.ToolStripMenuItem menuAddDate;
        private System.Windows.Forms.ToolStripMenuItem menuAppendDate;
        private System.Windows.Forms.ComboBox cmbDefaultFolder;
    }
}