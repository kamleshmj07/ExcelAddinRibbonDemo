namespace ExcelAddInDemo
{
    partial class VDRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public VDRibbon()
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
            this.tabMyVDTab = this.Factory.CreateRibbonTab();
            this.grpMyGroup = this.Factory.CreateRibbonGroup();
            this.ddlTable = this.Factory.CreateRibbonDropDown();
            this.groupInit = this.Factory.CreateRibbonGroup();
            this.btnConnect = this.Factory.CreateRibbonButton();
            this.btnDisconnect = this.Factory.CreateRibbonButton();
            this.tabMyVDTab.SuspendLayout();
            this.grpMyGroup.SuspendLayout();
            this.groupInit.SuspendLayout();
            this.SuspendLayout();
            // 
            // tabMyVDTab
            // 
            this.tabMyVDTab.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tabMyVDTab.Groups.Add(this.groupInit);
            this.tabMyVDTab.Groups.Add(this.grpMyGroup);
            this.tabMyVDTab.Label = "My VD Tab";
            this.tabMyVDTab.Name = "tabMyVDTab";
            this.tabMyVDTab.Position = this.Factory.RibbonPosition.AfterOfficeId("TabView");
            // 
            // grpMyGroup
            // 
            this.grpMyGroup.Items.Add(this.ddlTable);
            this.grpMyGroup.Label = "My Group";
            this.grpMyGroup.Name = "grpMyGroup";
            // 
            // ddlTable
            // 
            this.ddlTable.Label = "Tables";
            this.ddlTable.Name = "ddlTable";
            this.ddlTable.OfficeImageId = "TableInsert";
            this.ddlTable.ScreenTip = "Select Table";
            this.ddlTable.ShowImage = true;
            this.ddlTable.ShowLabel = false;
            this.ddlTable.SuperTip = "Allows to select a table";
            // 
            // groupInit
            // 
            this.groupInit.Items.Add(this.btnConnect);
            this.groupInit.Items.Add(this.btnDisconnect);
            this.groupInit.Label = "Init";
            this.groupInit.Name = "groupInit";
            // 
            // btnConnect
            // 
            this.btnConnect.Label = "Connect";
            this.btnConnect.Name = "btnConnect";
            this.btnConnect.OfficeImageId = "ServerConnection";
            this.btnConnect.ScreenTip = "Initiate Connection";
            this.btnConnect.ShowImage = true;
            this.btnConnect.SuperTip = "Click to connect to the database";
            // 
            // btnDisconnect
            // 
            this.btnDisconnect.Label = "Disconnect";
            this.btnDisconnect.Name = "btnDisconnect";
            this.btnDisconnect.OfficeImageId = "MailDelete";
            this.btnDisconnect.ScreenTip = "Close connection";
            this.btnDisconnect.ShowImage = true;
            this.btnDisconnect.SuperTip = "Click to close the connection";
            // 
            // VDRibbon
            // 
            this.Name = "VDRibbon";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tabMyVDTab);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.VDRibbon_Load);
            this.tabMyVDTab.ResumeLayout(false);
            this.tabMyVDTab.PerformLayout();
            this.grpMyGroup.ResumeLayout(false);
            this.grpMyGroup.PerformLayout();
            this.groupInit.ResumeLayout(false);
            this.groupInit.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tabMyVDTab;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpMyGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown ddlTable;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupInit;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnConnect;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnDisconnect;
    }

    partial class ThisRibbonCollection
    {
        internal VDRibbon VDRibbon
        {
            get { return this.GetRibbon<VDRibbon>(); }
        }
    }
}
