namespace MailCleanUp
{
    partial class MailCleanupRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public MailCleanupRibbon()
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MailCleanupRibbon));
            this.tab1 = this.Factory.CreateRibbonTab();
            this.MailCleanupGroup = this.Factory.CreateRibbonGroup();
            this.StartCleanupButton = this.Factory.CreateRibbonButton();
            this.label1 = this.Factory.CreateRibbonLabel();
            this.MaxItemsBox = this.Factory.CreateRibbonEditBox();
            this.tab1.SuspendLayout();
            this.MailCleanupGroup.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.MailCleanupGroup);
            this.tab1.Label = "TabAddIns";
            this.tab1.Name = "tab1";
            // 
            // MailCleanupGroup
            // 
            this.MailCleanupGroup.Items.Add(this.StartCleanupButton);
            this.MailCleanupGroup.Items.Add(this.label1);
            this.MailCleanupGroup.Items.Add(this.MaxItemsBox);
            this.MailCleanupGroup.Label = "Mail Cleanup";
            this.MailCleanupGroup.Name = "MailCleanupGroup";
            // 
            // StartCleanupButton
            // 
            this.StartCleanupButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.StartCleanupButton.Image = ((System.Drawing.Image)(resources.GetObject("StartCleanupButton.Image")));
            this.StartCleanupButton.Label = "Start Cleanup";
            this.StartCleanupButton.Name = "StartCleanupButton";
            this.StartCleanupButton.ShowImage = true;
            this.StartCleanupButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.StartCleanupButton_Click);
            // 
            // label1
            // 
            this.label1.Label = "Max items to delete:";
            this.label1.Name = "label1";
            // 
            // MaxItemsBox
            // 
            this.MaxItemsBox.Label = "editBox1";
            this.MaxItemsBox.MaxLength = 4;
            this.MaxItemsBox.Name = "MaxItemsBox";
            this.MaxItemsBox.ShowLabel = false;
            this.MaxItemsBox.TextChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.MaxItemsBox_TextChanged);
            // 
            // MailCleanupRibbon
            // 
            this.Name = "MailCleanupRibbon";
            this.RibbonType = "Microsoft.Outlook.Explorer";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.MailCleanupRibbon_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.MailCleanupGroup.ResumeLayout(false);
            this.MailCleanupGroup.PerformLayout();

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup MailCleanupGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton StartCleanupButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel label1;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox MaxItemsBox;
    }

    partial class ThisRibbonCollection
    {
        internal MailCleanupRibbon MailCleanupRibbon
        {
            get { return this.GetRibbon<MailCleanupRibbon>(); }
        }
    }
}
