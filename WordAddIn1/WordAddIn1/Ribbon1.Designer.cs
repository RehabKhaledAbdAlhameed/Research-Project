namespace WordAddIn1
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
            this.CompanyReport = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.companybtn = this.Factory.CreateRibbonButton();
            this.Publish = this.Factory.CreateRibbonTab();
            this.CompanyReport.SuspendLayout();
            this.group1.SuspendLayout();
            this.Publish.SuspendLayout();
            this.SuspendLayout();
            // 
            // CompanyReport
            // 
            this.CompanyReport.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.CompanyReport.Groups.Add(this.group1);
            this.CompanyReport.Label = "Company Report";
            this.CompanyReport.Name = "CompanyReport";
            // 
            // group1
            // 
            this.group1.Items.Add(this.companybtn);
            this.group1.Name = "group1";
            // 
            // companybtn
            // 
            this.companybtn.Label = "Reports";
            this.companybtn.Name = "companybtn";
            this.companybtn.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Company_Click);
            // 
            // Publish
            // 
            this.Publish.Label = "Publish";
            this.Publish.Name = "Publish";
            // 
            // Ribbon1
            // 
            this.Name = "Ribbon1";
            this.RibbonType = "Microsoft.Word.Document";
            this.Tabs.Add(this.CompanyReport);
            this.Tabs.Add(this.Publish);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.CompanyReport.ResumeLayout(false);
            this.CompanyReport.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.Publish.ResumeLayout(false);
            this.Publish.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab CompanyReport;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton companybtn;
        internal Microsoft.Office.Tools.Ribbon.RibbonTab Publish;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
