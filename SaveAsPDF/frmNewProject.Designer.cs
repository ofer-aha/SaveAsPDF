
namespace SaveAsPDF
{
    partial class frmNewProject
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmNewProject));
            this.txtProjectName = new System.Windows.Forms.TextBox();
            this.lblProjectName = new System.Windows.Forms.Label();
            this.lblProjectID = new System.Windows.Forms.Label();
            this.txtProjectId = new System.Windows.Forms.TextBox();
            this.tvDefaultSubFolders = new System.Windows.Forms.TreeView();
            this.menuTree = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.menueAdd = new System.Windows.Forms.ToolStripMenuItem();
            this.menuDel = new System.Windows.Forms.ToolStripMenuItem();
            this.menuRename = new System.Windows.Forms.ToolStripMenuItem();
            this.imageList = new System.Windows.Forms.ImageList(this.components);
            this.btmNewProject = new System.Windows.Forms.Button();
            this.txtProjectNotes = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.lblDefaultFoldersTree = new System.Windows.Forms.Label();
            this.btnCancel = new System.Windows.Forms.Button();
            this.menuTree.SuspendLayout();
            this.SuspendLayout();
            // 
            // txtProjectName
            // 
            this.txtProjectName.Location = new System.Drawing.Point(116, 81);
            this.txtProjectName.Name = "txtProjectName";
            this.txtProjectName.Size = new System.Drawing.Size(241, 20);
            this.txtProjectName.TabIndex = 1;
            // 
            // lblProjectName
            // 
            this.lblProjectName.AutoSize = true;
            this.lblProjectName.Location = new System.Drawing.Point(44, 81);
            this.lblProjectName.Name = "lblProjectName";
            this.lblProjectName.Size = new System.Drawing.Size(66, 13);
            this.lblProjectName.TabIndex = 1;
            this.lblProjectName.Text = "שם פרויקט:";
            this.lblProjectName.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // lblProjectID
            // 
            this.lblProjectID.AutoSize = true;
            this.lblProjectID.Location = new System.Drawing.Point(33, 46);
            this.lblProjectID.Name = "lblProjectID";
            this.lblProjectID.Size = new System.Drawing.Size(77, 13);
            this.lblProjectID.TabIndex = 2;
            this.lblProjectID.Text = "מספר פרויקט:";
            this.lblProjectID.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // txtProjectId
            // 
            this.txtProjectId.Location = new System.Drawing.Point(116, 46);
            this.txtProjectId.Name = "txtProjectId";
            this.txtProjectId.Size = new System.Drawing.Size(241, 20);
            this.txtProjectId.TabIndex = 0;
            // 
            // tvDefaultSubFolders
            // 
            this.tvDefaultSubFolders.ContextMenuStrip = this.menuTree;
            this.tvDefaultSubFolders.ImageIndex = 0;
            this.tvDefaultSubFolders.ImageList = this.imageList;
            this.tvDefaultSubFolders.Location = new System.Drawing.Point(507, 46);
            this.tvDefaultSubFolders.Name = "tvDefaultSubFolders";
            this.tvDefaultSubFolders.SelectedImageIndex = 0;
            this.tvDefaultSubFolders.Size = new System.Drawing.Size(224, 185);
            this.tvDefaultSubFolders.TabIndex = 5;
            this.tvDefaultSubFolders.AfterLabelEdit += new System.Windows.Forms.NodeLabelEditEventHandler(this.tvDefaultSubFolders_AfterLabelEdit);
            this.tvDefaultSubFolders.AfterSelect += new System.Windows.Forms.TreeViewEventHandler(this.tvDefaultSubFolders_AfterSelect);
            // 
            // menuTree
            // 
            this.menuTree.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.menueAdd,
            this.menuDel,
            this.menuRename});
            this.menuTree.Name = "cMenu_Add";
            this.menuTree.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.menuTree.Size = new System.Drawing.Size(138, 70);
            // 
            // menueAdd
            // 
            this.menueAdd.Name = "menueAdd";
            this.menueAdd.Size = new System.Drawing.Size(180, 22);
            this.menueAdd.Text = "הוסף תיקייה";
            this.menueAdd.Click += new System.EventHandler(this.menueAdd_Click);
            // 
            // menuDel
            // 
            this.menuDel.Name = "menuDel";
            this.menuDel.Size = new System.Drawing.Size(180, 22);
            this.menuDel.Text = "מחק תיקייה";
            this.menuDel.Click += new System.EventHandler(this.menuDel_Click);
            // 
            // menuRename
            // 
            this.menuRename.Name = "menuRename";
            this.menuRename.Size = new System.Drawing.Size(137, 22);
            this.menuRename.Text = "שנה שם";
            this.menuRename.Click += new System.EventHandler(this.menuRename_Click);
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
            // btmNewProject
            // 
            this.btmNewProject.Location = new System.Drawing.Point(261, 329);
            this.btmNewProject.Name = "btmNewProject";
            this.btmNewProject.Size = new System.Drawing.Size(126, 23);
            this.btmNewProject.TabIndex = 3;
            this.btmNewProject.Text = "&צור פרויקט חדש";
            this.btmNewProject.UseVisualStyleBackColor = true;
            this.btmNewProject.Click += new System.EventHandler(this.btmNewProject_Click);
            // 
            // txtProjectNotes
            // 
            this.txtProjectNotes.Location = new System.Drawing.Point(116, 116);
            this.txtProjectNotes.Multiline = true;
            this.txtProjectNotes.Name = "txtProjectNotes";
            this.txtProjectNotes.Size = new System.Drawing.Size(241, 115);
            this.txtProjectNotes.TabIndex = 2;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(21, 116);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(89, 13);
            this.label1.TabIndex = 10;
            this.label1.Text = "הערות לפרויקט:";
            this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // lblDefaultFoldersTree
            // 
            this.lblDefaultFoldersTree.AutoSize = true;
            this.lblDefaultFoldersTree.Location = new System.Drawing.Point(424, 46);
            this.lblDefaultFoldersTree.Name = "lblDefaultFoldersTree";
            this.lblDefaultFoldersTree.Size = new System.Drawing.Size(82, 13);
            this.lblDefaultFoldersTree.TabIndex = 11;
            this.lblDefaultFoldersTree.Text = "מבנה התיקיות:";
            this.lblDefaultFoldersTree.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // btnCancel
            // 
            this.btnCancel.Location = new System.Drawing.Point(402, 329);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(126, 23);
            this.btnCancel.TabIndex = 4;
            this.btnCancel.Text = "&ביטול";
            this.btnCancel.UseVisualStyleBackColor = true;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // frmNewProject
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(743, 364);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.lblDefaultFoldersTree);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.txtProjectNotes);
            this.Controls.Add(this.btmNewProject);
            this.Controls.Add(this.tvDefaultSubFolders);
            this.Controls.Add(this.txtProjectId);
            this.Controls.Add(this.lblProjectID);
            this.Controls.Add(this.lblProjectName);
            this.Controls.Add(this.txtProjectName);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "frmNewProject";
            this.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.RightToLeftLayout = true;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "פרויקט חדש";
            this.menuTree.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox txtProjectName;
        private System.Windows.Forms.Label lblProjectName;
        private System.Windows.Forms.Label lblProjectID;
        private System.Windows.Forms.TextBox txtProjectId;
        private System.Windows.Forms.TreeView tvDefaultSubFolders;
        private System.Windows.Forms.Button btmNewProject;
        private System.Windows.Forms.ImageList imageList;
        private System.Windows.Forms.ContextMenuStrip menuTree;
        private System.Windows.Forms.ToolStripMenuItem menueAdd;
        private System.Windows.Forms.ToolStripMenuItem menuDel;
        private System.Windows.Forms.ToolStripMenuItem menuRename;
        private System.Windows.Forms.TextBox txtProjectNotes;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label lblDefaultFoldersTree;
        private System.Windows.Forms.Button btnCancel;
    }
}