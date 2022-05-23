namespace SaveAsPDF
{
    partial class SaveAsPDFRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        private const string V = "Microsoft.Outlook.Mail.Compose, Microsoft.Outlook.Mail.Read, Microsoft.Outlook.Explorer";

        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public SaveAsPDFRibbon()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

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

        #region Component Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.SaveAsPDF = this.Factory.CreateRibbonTab();
            this.SaveAsPDFgrp = this.Factory.CreateRibbonGroup();
            this.buttonDemo = this.Factory.CreateRibbonButton();
            this.SaveAsPDF.SuspendLayout();
            this.SaveAsPDFgrp.SuspendLayout();
            this.SuspendLayout();
            // 
            // SaveAsPDF
            // 
            this.SaveAsPDF.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.SaveAsPDF.ControlId.OfficeId = "TabMail";
            this.SaveAsPDF.Groups.Add(this.SaveAsPDFgrp);
            this.SaveAsPDF.Label = "TabMail";
            this.SaveAsPDF.Name = "SaveAsPDF";
            // 
            // SaveAsPDFgrp
            // 
            this.SaveAsPDFgrp.Items.Add(this.buttonDemo);
            this.SaveAsPDFgrp.Label = "SaveAsPDF";
            this.SaveAsPDFgrp.Name = "SaveAsPDFgrp";
            // 
            // buttonDemo
            // 
            this.buttonDemo.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.buttonDemo.Image = global::SaveAsPDF.Properties.Resources.SaveAsPDF32;
            this.buttonDemo.Label = "SaveAsPDF";
            this.buttonDemo.Name = "buttonDemo";
            this.buttonDemo.ShowImage = true;
            this.buttonDemo.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonDemo_Click);
            // 
            // SaveAsPDFRibbon
            // 
            this.Name = "SaveAsPDFRibbon";
            this.RibbonType = "Microsoft.Outlook.Mail.Read, Microsoft.Outlook.Explorer";
            this.Tabs.Add(this.SaveAsPDF);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.SaveAsPDFRibbon_Load);
            this.SaveAsPDF.ResumeLayout(false);
            this.SaveAsPDF.PerformLayout();
            this.SaveAsPDFgrp.ResumeLayout(false);
            this.SaveAsPDFgrp.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab SaveAsPDF;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup SaveAsPDFgrp;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonDemo;
    }

    partial class ThisRibbonCollection
    {
        internal SaveAsPDFRibbon ribbSaveAsPDFRibbonon
        {
            get { return this.GetRibbon<SaveAsPDFRibbon>(); }
        }
    }
}
