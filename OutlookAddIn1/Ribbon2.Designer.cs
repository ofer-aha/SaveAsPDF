namespace SaveAsPDF1
{
    partial class MailItemRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public MailItemRibbon()
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
            this.Demo = this.Factory.CreateRibbonGroup();
            this.buttonDemo = this.Factory.CreateRibbonButton();
            this.button1 = this.Factory.CreateRibbonButton();
            this.button2 = this.Factory.CreateRibbonButton();
            this.SaveAsPDF.SuspendLayout();
            this.Demo.SuspendLayout();
            this.SuspendLayout();
            // 
            // SaveAsPDF
            // 
            this.SaveAsPDF.Groups.Add(this.Demo);
            this.SaveAsPDF.Label = "SaveAsPDF";
            this.SaveAsPDF.Name = "SaveAsPDF";
            // 
            // Demo
            // 
            this.Demo.Items.Add(this.buttonDemo);
            this.Demo.Items.Add(this.button1);
            this.Demo.Items.Add(this.button2);
            this.Demo.Label = "MailItem";
            this.Demo.Name = "Demo";
            // 
            // buttonDemo
            // 
            this.buttonDemo.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.buttonDemo.Image = global::SaveAsPDF1.Properties.Resources.arrow2;
            this.buttonDemo.Label = "ButtonDemo";
            this.buttonDemo.Name = "buttonDemo";
            this.buttonDemo.ShowImage = true;
            this.buttonDemo.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonDemo_Click);
            // 
            // button1
            // 
            this.button1.Label = "button2Demo";
            this.button1.Name = "button1";
            this.button1.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button1_Click);
            // 
            // button2
            // 
            this.button2.Label = "button2";
            this.button2.Name = "button2";
            this.button2.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button2_Click);
            // 
            // MailItemRibbon
            // 
            this.Name = "MailItemRibbon";
            this.RibbonType = "Microsoft.Outlook.Mail.Compose, Microsoft.Outlook.Mail.Read";
            this.Tabs.Add(this.SaveAsPDF);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon2_Load);
            this.SaveAsPDF.ResumeLayout(false);
            this.SaveAsPDF.PerformLayout();
            this.Demo.ResumeLayout(false);
            this.Demo.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab SaveAsPDF;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup Demo;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonDemo;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button2;
    }

    partial class ThisRibbonCollection
    {
        internal MailItemRibbon Ribbon2
        {
            get { return this.GetRibbon<MailItemRibbon>(); }
        }
    }
}
