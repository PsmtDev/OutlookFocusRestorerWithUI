namespace OutlookFocusRestorerWithUI
{
    partial class FocusRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        private System.ComponentModel.IContainer components = null;

        public FocusRibbon()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing && components != null)
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        private void InitializeComponent()
        {
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.btnRestoreLastMail = this.Factory.CreateRibbonButton();
            this.tab1.Label = "Tab1";
            this.tab1.Groups.Add(this.group1);
            this.group1.Label = "Focus Tools";
            this.group1.Items.Add(this.btnRestoreLastMail);
            this.btnRestoreLastMail.Label = "Restaurar último correo";
            this.btnRestoreLastMail.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnRestoreLastMail_Click);
            this.Name = "FocusRibbon";
            this.RibbonType = "Microsoft.Outlook.Explorer";
            this.Tabs.Add(this.tab1);
        }

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnRestoreLastMail;
    }

    partial class ThisRibbonCollection
    {
        internal FocusRibbon FocusRibbon
        {
            get { return this.GetRibbon<FocusRibbon>(); }
        }
    }
}
