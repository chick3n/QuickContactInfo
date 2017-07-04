namespace QuickContactInfo
{
    partial class Ribbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon()
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
            this.tabView = this.Factory.CreateRibbonTab();
            this.grpToggle = this.Factory.CreateRibbonGroup();
            this.mnuToggle = this.Factory.CreateRibbonMenu();
            this.btnEnableMessage = this.Factory.CreateRibbonToggleButton();
            this.btnEnableMeeting = this.Factory.CreateRibbonToggleButton();
            this.btnOff = this.Factory.CreateRibbonToggleButton();
            this.tabView.SuspendLayout();
            this.grpToggle.SuspendLayout();
            this.SuspendLayout();
            // 
            // tabView
            // 
            this.tabView.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tabView.ControlId.OfficeId = "TabView";
            this.tabView.Groups.Add(this.grpToggle);
            this.tabView.Label = "TabView";
            this.tabView.Name = "tabView";
            // 
            // grpToggle
            // 
            this.grpToggle.Items.Add(this.mnuToggle);
            this.grpToggle.Label = "Quick Pane";
            this.grpToggle.Name = "grpToggle";
            // 
            // mnuToggle
            // 
            this.mnuToggle.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.mnuToggle.Image = global::QuickContactInfo.Properties.Resources.RibbonIcon;
            this.mnuToggle.Items.Add(this.btnEnableMessage);
            this.mnuToggle.Items.Add(this.btnEnableMeeting);
            this.mnuToggle.Items.Add(this.btnOff);
            this.mnuToggle.Label = "Quick Pane";
            this.mnuToggle.Name = "mnuToggle";
            this.mnuToggle.ShowImage = true;
            // 
            // btnEnableMessage
            // 
            this.btnEnableMessage.Label = "Messages";
            this.btnEnableMessage.Name = "btnEnableMessage";
            this.btnEnableMessage.ShowImage = true;
            this.btnEnableMessage.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnEnableMessage_Click);
            // 
            // btnEnableMeeting
            // 
            this.btnEnableMeeting.Label = "Meetings";
            this.btnEnableMeeting.Name = "btnEnableMeeting";
            this.btnEnableMeeting.ShowImage = true;
            this.btnEnableMeeting.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnEnableMeeting_Click);
            // 
            // btnOff
            // 
            this.btnOff.Label = "Off";
            this.btnOff.Name = "btnOff";
            this.btnOff.ShowImage = true;
            this.btnOff.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnOff_Click);
            // 
            // Ribbon
            // 
            this.Name = "Ribbon";
            this.RibbonType = "Microsoft.Outlook.Explorer";
            this.Tabs.Add(this.tabView);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon_Load);
            this.tabView.ResumeLayout(false);
            this.tabView.PerformLayout();
            this.grpToggle.ResumeLayout(false);
            this.grpToggle.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tabView;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpToggle;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu mnuToggle;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton btnEnableMessage;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton btnEnableMeeting;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton btnOff;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon Ribbon
        {
            get { return this.GetRibbon<Ribbon>(); }
        }
    }
}
