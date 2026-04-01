
namespace SaveAsPDF
{
    partial class FormContacts
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FormContacts));
            this.splitContainer = new System.Windows.Forms.SplitContainer();
            this.splitContainerInner = new System.Windows.Forms.SplitContainer();
            this.lblLoading = new System.Windows.Forms.Label();
            this.dgvContacs = new System.Windows.Forms.DataGridView();
            this.tvFolders = new System.Windows.Forms.TreeView();
            this.btnCancel = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.txtFilter = new System.Windows.Forms.TextBox();
            this.btnOK = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer)).BeginInit();
            this.splitContainer.Panel1.SuspendLayout();
            this.splitContainer.Panel2.SuspendLayout();
            this.splitContainer.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.splitContainerInner)).BeginInit();
            this.splitContainerInner.Panel1.SuspendLayout();
            this.splitContainerInner.Panel2.SuspendLayout();
            this.splitContainerInner.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvContacs)).BeginInit();
            this.SuspendLayout();
            // 
            // splitContainer
            // 
            this.splitContainer.Dock = System.Windows.Forms.DockStyle.Fill;
            this.splitContainer.Location = new System.Drawing.Point(0, 0);
            this.splitContainer.Name = "splitContainer";
            this.splitContainer.Orientation = System.Windows.Forms.Orientation.Horizontal;
            // 
            // splitContainer.Panel1
            // 
            this.splitContainer.Panel1.Controls.Add(this.splitContainerInner);
            this.splitContainer.Panel1.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            // 
            // splitContainer.Panel2
            // 
            this.splitContainer.Panel2.Controls.Add(this.btnCancel);
            this.splitContainer.Panel2.Controls.Add(this.label2);
            this.splitContainer.Panel2.Controls.Add(this.label1);
            this.splitContainer.Panel2.Controls.Add(this.txtFilter);
            this.splitContainer.Panel2.Controls.Add(this.btnOK);
            this.splitContainer.Panel2.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.splitContainer.Size = new System.Drawing.Size(664, 418);
            this.splitContainer.SplitterDistance = 353;
            this.splitContainer.TabIndex = 0;
            // 
            // splitContainerInner
            // 
            this.splitContainerInner.Dock = System.Windows.Forms.DockStyle.Fill;
            this.splitContainerInner.Location = new System.Drawing.Point(0, 0);
            this.splitContainerInner.Name = "splitContainerInner";
            // 
            // splitContainerInner.Panel1 - DataGridView (left / main area)
            // 
            this.splitContainerInner.Panel1.Controls.Add(this.lblLoading);
            this.splitContainerInner.Panel1.Controls.Add(this.dgvContacs);
            // 
            // splitContainerInner.Panel2 - TreeView (right / folder navigation for RTL)
            // 
            this.splitContainerInner.Panel2.Controls.Add(this.tvFolders);
            this.splitContainerInner.Size = new System.Drawing.Size(664, 353);
            this.splitContainerInner.SplitterDistance = 464;
            this.splitContainerInner.TabIndex = 0;
            // 
            // lblLoading
            // 
            this.lblLoading.AutoSize = true;
            this.lblLoading.Font = new System.Drawing.Font("Microsoft Sans Serif", 13F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblLoading.Location = new System.Drawing.Point(138, 130);
            this.lblLoading.Name = "lblLoading";
            this.lblLoading.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.lblLoading.Size = new System.Drawing.Size(134, 22);
            this.lblLoading.TabIndex = 1;
            this.lblLoading.Text = ".......טוען נתונים";
            // 
            // dgvContacs
            // 
            this.dgvContacs.AllowUserToAddRows = false;
            this.dgvContacs.AllowUserToOrderColumns = true;
            this.dgvContacs.AllowUserToResizeRows = false;
            this.dgvContacs.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells;
            this.dgvContacs.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells;
            this.dgvContacs.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvContacs.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dgvContacs.Location = new System.Drawing.Point(0, 0);
            this.dgvContacs.MultiSelect = false;
            this.dgvContacs.Name = "dgvContacs";
            this.dgvContacs.RowHeadersVisible = false;
            this.dgvContacs.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dgvContacs.Size = new System.Drawing.Size(464, 353);
            this.dgvContacs.TabIndex = 0;
            this.dgvContacs.DoubleClick += new System.EventHandler(this.dgvContacs_DoubleClick);
            // 
            // tvFolders
            // 
            this.tvFolders.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tvFolders.Location = new System.Drawing.Point(0, 0);
            this.tvFolders.Name = "tvFolders";
            this.tvFolders.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.tvFolders.RightToLeftLayout = true;
            this.tvFolders.Size = new System.Drawing.Size(196, 353);
            this.tvFolders.TabIndex = 0;
            this.tvFolders.AfterSelect += new System.Windows.Forms.TreeViewEventHandler(this.tvFolders_AfterSelect);
            // 
            // btnCancel
            // 
            this.btnCancel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnCancel.Location = new System.Drawing.Point(458, 20);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(94, 24);
            this.btnCancel.TabIndex = 4;
            this.btnCancel.Text = "&ביטול";
            this.btnCancel.UseVisualStyleBackColor = true;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // label2
            // 
            this.label2.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(351, 4);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(301, 13);
            this.label2.TabIndex = 3;
            this.label2.Text = "שם לב: אנשי קשר ללא כתובת אי מייל לא יופיעו ברשימה";
            // 
            // label1
            // 
            this.label1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(195, 26);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(40, 13);
            this.label1.TabIndex = 2;
            this.label1.Text = "חיפוש";
            // 
            // txtFilter
            // 
            this.txtFilter.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.txtFilter.Location = new System.Drawing.Point(13, 22);
            this.txtFilter.Name = "txtFilter";
            this.txtFilter.Size = new System.Drawing.Size(176, 20);
            this.txtFilter.TabIndex = 1;
            this.txtFilter.TextChanged += new System.EventHandler(this.txtFilter_TextChanged);
            // 
            // btnOK
            // 
            this.btnOK.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnOK.Location = new System.Drawing.Point(558, 20);
            this.btnOK.Name = "btnOK";
            this.btnOK.Size = new System.Drawing.Size(94, 24);
            this.btnOK.TabIndex = 0;
            this.btnOK.Text = "&הוסף";
            this.btnOK.UseVisualStyleBackColor = true;
            this.btnOK.Click += new System.EventHandler(this.btnOK_Click);
            // 
            // FormContacts
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(664, 418);
            this.Controls.Add(this.splitContainer);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "FormContacts";
            this.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "בחר אנשי קשר";
            this.Activated += new System.EventHandler(this.FormContacts_Activated);
            this.Load += new System.EventHandler(this.FormContacts_Load);
            this.splitContainer.Panel1.ResumeLayout(false);
            this.splitContainer.Panel2.ResumeLayout(false);
            this.splitContainer.Panel2.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer)).EndInit();
            this.splitContainer.ResumeLayout(false);
            this.splitContainerInner.Panel1.ResumeLayout(false);
            this.splitContainerInner.Panel1.PerformLayout();
            this.splitContainerInner.Panel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.splitContainerInner)).EndInit();
            this.splitContainerInner.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dgvContacs)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.SplitContainer splitContainer;
        private System.Windows.Forms.SplitContainer splitContainerInner;
        private System.Windows.Forms.DataGridView dgvContacs;
        private System.Windows.Forms.TreeView tvFolders;
        private System.Windows.Forms.Button btnOK;
        private System.Windows.Forms.TextBox txtFilter;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label lblLoading;
        private System.Windows.Forms.Button btnCancel;
    }
}