namespace OutlookAutomation
{
    partial class CustomRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public CustomRibbon()
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
            this.CustomTab = this.Factory.CreateRibbonTab();
            this.automationGroup = this.Factory.CreateRibbonGroup();
            this.printPane = this.Factory.CreateRibbonButton();
            this.CustomTab.SuspendLayout();
            this.automationGroup.SuspendLayout();
            this.SuspendLayout();
            // 
            // CustomTab
            // 
            this.CustomTab.Groups.Add(this.automationGroup);
            this.CustomTab.Label = "Custom Tab";
            this.CustomTab.Name = "CustomTab";
            // 
            // automationGroup
            // 
            this.automationGroup.Items.Add(this.printPane);
            this.automationGroup.Label = "Automation";
            this.automationGroup.Name = "automationGroup";
            // 
            // printPane
            // 
            this.printPane.Label = "Export";
            this.printPane.Name = "printPane";
            this.printPane.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.printPane_Click);
            // 
            // CustomRibbon
            // 
            this.Name = "CustomRibbon";
            this.RibbonType = "Microsoft.Outlook.Explorer";
            this.Tabs.Add(this.CustomTab);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.CustomTab.ResumeLayout(false);
            this.CustomTab.PerformLayout();
            this.automationGroup.ResumeLayout(false);
            this.automationGroup.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion
        internal Microsoft.Office.Tools.Ribbon.RibbonTab CustomTab;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup automationGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton printPane;
    }

    partial class ThisRibbonCollection
    {
        internal CustomRibbon Ribbon1
        {
            get { return this.GetRibbon<CustomRibbon>(); }
        }
    }
}
