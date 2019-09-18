namespace BoschTool
{
    partial class Ribbon1 : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon1()
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
            this.tabBosch = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.button3 = this.Factory.CreateRibbonButton();
            this.button1 = this.Factory.CreateRibbonButton();
            this.button6 = this.Factory.CreateRibbonButton();
            this.button4 = this.Factory.CreateRibbonButton();
            this.button5 = this.Factory.CreateRibbonButton();
            this.button2 = this.Factory.CreateRibbonButton();
            this.tabBosch.SuspendLayout();
            this.group1.SuspendLayout();
            this.SuspendLayout();
            // 
            // tabBosch
            // 
            this.tabBosch.Groups.Add(this.group1);
            this.tabBosch.Label = "BOSCH";
            this.tabBosch.Name = "tabBosch";
            // 
            // group1
            // 
            this.group1.Items.Add(this.button3);
            this.group1.Items.Add(this.button1);
            this.group1.Items.Add(this.button6);
            this.group1.Items.Add(this.button4);
            this.group1.Items.Add(this.button5);
            this.group1.Items.Add(this.button2);
            this.group1.Label = "Web Functions";
            this.group1.Name = "group1";
            // 
            // button3
            // 
            this.button3.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button3.Label = "Log In";
            this.button3.Name = "button3";
            this.button3.OfficeImageId = "CondolatoryEvent";
            this.button3.ShowImage = true;
            this.button3.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Button3_Click);
            // 
            // button1
            // 
            this.button1.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button1.Label = "Get File List";
            this.button1.Name = "button1";
            this.button1.OfficeImageId = "AdpDiagramColumnNames";
            this.button1.ShowImage = true;
            this.button1.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button1_Click);
            // 
            // button6
            // 
            this.button6.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button6.Label = "Upload Invoice";
            this.button6.Name = "button6";
            this.button6.OfficeImageId = "DatabaseSqlServer";
            this.button6.ShowImage = true;
            this.button6.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Button6_Click);
            // 
            // button4
            // 
            this.button4.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button4.Label = "Assign Shipment";
            this.button4.Name = "button4";
            this.button4.OfficeImageId = "AdpDiagramAddRelatedTables";
            this.button4.ShowImage = true;
            this.button4.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Button4_Click);
            // 
            // button5
            // 
            this.button5.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button5.Label = "Generate CCS";
            this.button5.Name = "button5";
            this.button5.OfficeImageId = "CreateEmail";
            this.button5.ShowImage = true;
            this.button5.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Button5_Click);
            // 
            // button2
            // 
            this.button2.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button2.Label = "Donwload Docs";
            this.button2.Name = "button2";
            this.button2.OfficeImageId = "SaveSentItemsMenu";
            this.button2.ShowImage = true;
            this.button2.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button2_Click);
            // 
            // Ribbon1
            // 
            this.Name = "Ribbon1";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tabBosch);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.tabBosch.ResumeLayout(false);
            this.tabBosch.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        public Microsoft.Office.Tools.Ribbon.RibbonTab tabBosch;
        public Microsoft.Office.Tools.Ribbon.RibbonButton button1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button3;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button6;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button4;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button5;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
