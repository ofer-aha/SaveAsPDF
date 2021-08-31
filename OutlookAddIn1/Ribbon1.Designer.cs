namespace SaveAsPDF1
{
    partial class ExlorerRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public ExlorerRibbon()
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
            this.components = new System.ComponentModel.Container();
            this.SaveAsPDF = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.button1 = this.Factory.CreateRibbonButton();
            this.button2 = this.Factory.CreateRibbonButton();
            this.imageList1 = new System.Windows.Forms.ImageList(this.components);
            this.SaveAsPDF.SuspendLayout();
            this.group1.SuspendLayout();
            this.SuspendLayout();
            // 
            // SaveAsPDF
            // 
            this.SaveAsPDF.Groups.Add(this.group1);
            this.SaveAsPDF.Label = "SaveAsPDF";
            this.SaveAsPDF.Name = "SaveAsPDF";
            // 
            // group1
            // 
            this.group1.Items.Add(this.button1);
            this.group1.Label = "Explorer";
            this.group1.Name = "group1";
            // 
            // button1
            // 
            this.button1.Label = "SaveAsPDF";
            this.button1.Name = "button1";
            this.button1.ShowImage = true;
            this.button1.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button1_Click);
            // 
            // button2
            // 
            this.button2.Label = "button2";
            this.button2.Name = "button2";
            this.button2.ShowImage = true;
            // 
            // imageList1
            // 
            this.imageList1.ColorDepth = System.Windows.Forms.ColorDepth.Depth8Bit;
            this.imageList1.ImageSize = new System.Drawing.Size(16, 16);
            this.imageList1.TransparentColor = System.Drawing.Color.Transparent;
            // 
            // ExlorerRibbon
            // 
            this.Name = "ExlorerRibbon";
            // 
            // ExlorerRibbon.OfficeMenu
            // 
            this.OfficeMenu.Items.Add(this.button2);
            this.RibbonType = "Microsoft.Outlook.Explorer";
            this.Tabs.Add(this.SaveAsPDF);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.SaveAsPDF.ResumeLayout(false);
            this.SaveAsPDF.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab SaveAsPDF;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button2;
        private System.Windows.Forms.ImageList imageList1;
    }

    partial class ThisRibbonCollection
    {
        internal ExlorerRibbon Ribbon1
        {
            get { return this.GetRibbon<ExlorerRibbon>(); }
        }
    }
}
