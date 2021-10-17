
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmNewProject));
            this.txtProjectName = new System.Windows.Forms.TextBox();
            this.lblProjectName = new System.Windows.Forms.Label();
            this.lblProjectID = new System.Windows.Forms.Label();
            this.txtProjectId = new System.Windows.Forms.TextBox();
            this.dgvDefaultSubFolders = new System.Windows.Forms.DataGridView();
            this.tvDefaultSubFolders = new System.Windows.Forms.TreeView();
            this.btmNewProject = new System.Windows.Forms.Button();
            this.btnAddNode = new System.Windows.Forms.Button();
            this.txtNode = new System.Windows.Forms.TextBox();
            ((System.ComponentModel.ISupportInitialize)(this.dgvDefaultSubFolders)).BeginInit();
            this.SuspendLayout();
            // 
            // txtProjectName
            // 
            this.txtProjectName.Location = new System.Drawing.Point(391, 74);
            this.txtProjectName.Name = "txtProjectName";
            this.txtProjectName.Size = new System.Drawing.Size(241, 20);
            this.txtProjectName.TabIndex = 0;
            this.txtProjectName.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            // 
            // lblProjectName
            // 
            this.lblProjectName.AutoSize = true;
            this.lblProjectName.Location = new System.Drawing.Point(664, 74);
            this.lblProjectName.Name = "lblProjectName";
            this.lblProjectName.Size = new System.Drawing.Size(66, 13);
            this.lblProjectName.TabIndex = 1;
            this.lblProjectName.Text = "שם פרויקט:";
            this.lblProjectName.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // lblProjectID
            // 
            this.lblProjectID.AutoSize = true;
            this.lblProjectID.Location = new System.Drawing.Point(653, 45);
            this.lblProjectID.Name = "lblProjectID";
            this.lblProjectID.Size = new System.Drawing.Size(77, 13);
            this.lblProjectID.TabIndex = 2;
            this.lblProjectID.Text = "מספר פרויקט:";
            this.lblProjectID.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // txtProjectId
            // 
            this.txtProjectId.Location = new System.Drawing.Point(391, 38);
            this.txtProjectId.Name = "txtProjectId";
            this.txtProjectId.Size = new System.Drawing.Size(241, 20);
            this.txtProjectId.TabIndex = 3;
            this.txtProjectId.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            // 
            // dgvDefaultSubFolders
            // 
            this.dgvDefaultSubFolders.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvDefaultSubFolders.Location = new System.Drawing.Point(391, 136);
            this.dgvDefaultSubFolders.Name = "dgvDefaultSubFolders";
            this.dgvDefaultSubFolders.Size = new System.Drawing.Size(241, 150);
            this.dgvDefaultSubFolders.TabIndex = 4;
            // 
            // tvDefaultSubFolders
            // 
            this.tvDefaultSubFolders.Location = new System.Drawing.Point(70, 136);
            this.tvDefaultSubFolders.Name = "tvDefaultSubFolders";
            this.tvDefaultSubFolders.Size = new System.Drawing.Size(224, 150);
            this.tvDefaultSubFolders.TabIndex = 5;
            // 
            // btmNewProject
            // 
            this.btmNewProject.Location = new System.Drawing.Point(296, 367);
            this.btmNewProject.Name = "btmNewProject";
            this.btmNewProject.Size = new System.Drawing.Size(126, 23);
            this.btmNewProject.TabIndex = 6;
            this.btmNewProject.Text = "צור פרויקט חדש";
            this.btmNewProject.UseVisualStyleBackColor = true;
            this.btmNewProject.Click += new System.EventHandler(this.btmNewProject_Click);
            // 
            // btnAddNode
            // 
            this.btnAddNode.Location = new System.Drawing.Point(70, 293);
            this.btnAddNode.Name = "btnAddNode";
            this.btnAddNode.Size = new System.Drawing.Size(75, 23);
            this.btnAddNode.TabIndex = 7;
            this.btnAddNode.Text = "node";
            this.btnAddNode.UseVisualStyleBackColor = true;
            this.btnAddNode.Click += new System.EventHandler(this.btnAddNode_Click);
            // 
            // txtNode
            // 
            this.txtNode.Location = new System.Drawing.Point(70, 110);
            this.txtNode.Name = "txtNode";
            this.txtNode.Size = new System.Drawing.Size(224, 20);
            this.txtNode.TabIndex = 8;
            this.txtNode.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            // 
            // frmNewProject
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.txtNode);
            this.Controls.Add(this.btnAddNode);
            this.Controls.Add(this.btmNewProject);
            this.Controls.Add(this.tvDefaultSubFolders);
            this.Controls.Add(this.dgvDefaultSubFolders);
            this.Controls.Add(this.txtProjectId);
            this.Controls.Add(this.lblProjectID);
            this.Controls.Add(this.lblProjectName);
            this.Controls.Add(this.txtProjectName);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "frmNewProject";
            this.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.Text = "פרויקט חדש";
            ((System.ComponentModel.ISupportInitialize)(this.dgvDefaultSubFolders)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox txtProjectName;
        private System.Windows.Forms.Label lblProjectName;
        private System.Windows.Forms.Label lblProjectID;
        private System.Windows.Forms.TextBox txtProjectId;
        private System.Windows.Forms.DataGridView dgvDefaultSubFolders;
        private System.Windows.Forms.TreeView tvDefaultSubFolders;
        private System.Windows.Forms.Button btmNewProject;
        private System.Windows.Forms.Button btnAddNode;
        private System.Windows.Forms.TextBox txtNode;
    }
}