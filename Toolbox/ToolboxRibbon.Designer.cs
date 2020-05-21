namespace Toolbox
{
    partial class ToolboxRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public ToolboxRibbon()
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
            this.tabToolbox = this.Factory.CreateRibbonTab();
            this.grpUnits = this.Factory.CreateRibbonGroup();
            this.btnFixQuantity = this.Factory.CreateRibbonButton();
            this.tabToolboxTab = this.Factory.CreateRibbonTab();
            this.tabToolbox.SuspendLayout();
            this.grpUnits.SuspendLayout();
            this.tabToolboxTab.SuspendLayout();
            this.SuspendLayout();
            // 
            // tabToolbox
            // 
            this.tabToolbox.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tabToolbox.ControlId.OfficeId = "TabAddins";
            this.tabToolbox.Groups.Add(this.grpUnits);
            this.tabToolbox.Label = "Toolbox";
            this.tabToolbox.Name = "tabToolbox";
            // 
            // grpUnits
            // 
            this.grpUnits.Items.Add(this.btnFixQuantity);
            this.grpUnits.Label = "Units";
            this.grpUnits.Name = "grpUnits";
            // 
            // btnFixQuantity
            // 
            this.btnFixQuantity.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnFixQuantity.KeyTip = "Q";
            this.btnFixQuantity.Label = "Fix Quantities";
            this.btnFixQuantity.Name = "btnFixQuantity";
            this.btnFixQuantity.OfficeImageId = "AutoFormatWizard";
            this.btnFixQuantity.ScreenTip = "Fix Quantities";
            this.btnFixQuantity.ShowImage = true;
            this.btnFixQuantity.SuperTip = "Scans the document and corrects quantities with attached units to conform to NIST" +
    " Special Publication 811. E.g., will convert \"1+/-0.01mA\" to \"1.00 mA ± 0.01 mA\"" +
    "";
            this.btnFixQuantity.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnFixQuantity_Click);
            // 
            // tabToolboxTab
            // 
            this.tabToolboxTab.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tabToolboxTab.ControlId.OfficeId = "TabToolbox";
            this.tabToolboxTab.Label = "Toolbox";
            this.tabToolboxTab.Name = "tabToolboxTab";
            // 
            // ToolboxRibbon
            // 
            this.Name = "ToolboxRibbon";
            this.RibbonType = "Microsoft.Word.Document";
            this.Tabs.Add(this.tabToolbox);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Toolbox_Load);
            this.tabToolbox.ResumeLayout(false);
            this.tabToolbox.PerformLayout();
            this.grpUnits.ResumeLayout(false);
            this.grpUnits.PerformLayout();
            this.tabToolboxTab.ResumeLayout(false);
            this.tabToolboxTab.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tabToolbox;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpUnits;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnFixQuantity;
        internal Microsoft.Office.Tools.Ribbon.RibbonTab tabToolboxTab;
    }

    partial class ThisRibbonCollection
    {
        internal ToolboxRibbon Toolbox
        {
            get { return this.GetRibbon<ToolboxRibbon>(); }
        }
    }
}
