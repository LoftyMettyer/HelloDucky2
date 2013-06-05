namespace Fusion
{
	partial class MainForm
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

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

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
			this.components = new System.ComponentModel.Container();
			Infragistics.Win.UltraWinToolbars.OptionSet optionSet1 = new Infragistics.Win.UltraWinToolbars.OptionSet("Show");
			Infragistics.Win.UltraWinToolbars.RibbonTab ribbonTab1 = new Infragistics.Win.UltraWinToolbars.RibbonTab("Config");
			Infragistics.Win.UltraWinToolbars.RibbonGroup ribbonGroup1 = new Infragistics.Win.UltraWinToolbars.RibbonGroup("ribbonGroup1");
			Infragistics.Win.UltraWinToolbars.ButtonTool buttonTool1 = new Infragistics.Win.UltraWinToolbars.ButtonTool("ID_Save");
			Infragistics.Win.UltraWinToolbars.RibbonGroup ribbonGroup2 = new Infragistics.Win.UltraWinToolbars.RibbonGroup("ribbonGroup2");
			Infragistics.Win.UltraWinToolbars.StateButtonTool stateButtonTool1 = new Infragistics.Win.UltraWinToolbars.StateButtonTool("ID_ShowConfiguration", "Show");
			Infragistics.Win.UltraWinToolbars.StateButtonTool stateButtonTool2 = new Infragistics.Win.UltraWinToolbars.StateButtonTool("ID_ShowMessages", "Show");
			Infragistics.Win.UltraWinToolbars.StateButtonTool stateButtonTool3 = new Infragistics.Win.UltraWinToolbars.StateButtonTool("ID_ShowLogs", "Show");
			Infragistics.Win.UltraWinToolbars.RibbonGroup ribbonGroup3 = new Infragistics.Win.UltraWinToolbars.RibbonGroup("ribbonGroup3");
			Infragistics.Win.UltraWinToolbars.ButtonTool buttonTool3 = new Infragistics.Win.UltraWinToolbars.ButtonTool("ID_Purge");
			Infragistics.Win.UltraWinToolbars.ButtonTool buttonTool5 = new Infragistics.Win.UltraWinToolbars.ButtonTool("ID_LogExport");
			Infragistics.Win.UltraWinToolbars.RibbonTab ribbonTab2 = new Infragistics.Win.UltraWinToolbars.RibbonTab("SystemSettings");
			Infragistics.Win.UltraWinToolbars.RibbonGroup ribbonGroup4 = new Infragistics.Win.UltraWinToolbars.RibbonGroup("ribbonGroup1");
			Infragistics.Win.UltraWinToolbars.ButtonTool buttonTool14 = new Infragistics.Win.UltraWinToolbars.ButtonTool("ID_Save");
			Infragistics.Win.UltraWinToolbars.RibbonGroup ribbonGroup5 = new Infragistics.Win.UltraWinToolbars.RibbonGroup("ribbonGroup2");
			Infragistics.Win.UltraWinToolbars.ButtonTool buttonTool9 = new Infragistics.Win.UltraWinToolbars.ButtonTool("ID_StartConnector");
			Infragistics.Win.UltraWinToolbars.ButtonTool buttonTool13 = new Infragistics.Win.UltraWinToolbars.ButtonTool("ID_StopConnector");
			Infragistics.Win.UltraWinToolbars.ButtonTool buttonTool2 = new Infragistics.Win.UltraWinToolbars.ButtonTool("ID_Save");
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinToolbars.StateButtonTool stateButtonTool4 = new Infragistics.Win.UltraWinToolbars.StateButtonTool("ID_ShowConfiguration", "Show");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinToolbars.StateButtonTool stateButtonTool5 = new Infragistics.Win.UltraWinToolbars.StateButtonTool("ID_ShowMessages", "Show");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinToolbars.StateButtonTool stateButtonTool6 = new Infragistics.Win.UltraWinToolbars.StateButtonTool("ID_ShowLogs", "Show");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinToolbars.ButtonTool buttonTool4 = new Infragistics.Win.UltraWinToolbars.ButtonTool("ID_Purge");
			Infragistics.Win.UltraWinToolbars.ButtonTool buttonTool6 = new Infragistics.Win.UltraWinToolbars.ButtonTool("ID_LogExport");
			Infragistics.Win.UltraWinToolbars.ButtonTool buttonTool11 = new Infragistics.Win.UltraWinToolbars.ButtonTool("ID_StartConnector");
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinToolbars.ButtonTool buttonTool12 = new Infragistics.Win.UltraWinToolbars.ButtonTool("ID_StopConnector");
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			this.mainForm_Fill_Panel = new Infragistics.Win.Misc.UltraPanel();
			this.viewPanel = new System.Windows.Forms.Panel();
			this._mainForm_Toolbars_Dock_Area_Left = new Infragistics.Win.UltraWinToolbars.UltraToolbarsDockArea();
			this._mainForm_Toolbars_Dock_Area_Right = new Infragistics.Win.UltraWinToolbars.UltraToolbarsDockArea();
			this._mainForm_Toolbars_Dock_Area_Top = new Infragistics.Win.UltraWinToolbars.UltraToolbarsDockArea();
			this._mainForm_Toolbars_Dock_Area_Bottom = new Infragistics.Win.UltraWinToolbars.UltraToolbarsDockArea();
			this.timer1 = new System.Windows.Forms.Timer(this.components);
			this.pageConfiguration = new Fusion.Pages.Configuration();
			this.pageSettings = new Fusion.Pages.Settings();
			this.pageLogs = new Fusion.Pages.Logs();
			this.pageMessages = new Fusion.Pages.Messages();
			this.toolbarsManager = new Infragistics.Win.UltraWinToolbars.UltraToolbarsManager(this.components);
			this.mainForm_Fill_Panel.ClientArea.SuspendLayout();
			this.mainForm_Fill_Panel.SuspendLayout();
			this.viewPanel.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.toolbarsManager)).BeginInit();
			this.SuspendLayout();
			// 
			// mainForm_Fill_Panel
			// 
			// 
			// mainForm_Fill_Panel.ClientArea
			// 
			this.mainForm_Fill_Panel.ClientArea.Controls.Add(this.viewPanel);
			this.mainForm_Fill_Panel.Cursor = System.Windows.Forms.Cursors.Default;
			this.mainForm_Fill_Panel.Dock = System.Windows.Forms.DockStyle.Fill;
			this.mainForm_Fill_Panel.Location = new System.Drawing.Point(0, 127);
			this.mainForm_Fill_Panel.Name = "mainForm_Fill_Panel";
			this.mainForm_Fill_Panel.Size = new System.Drawing.Size(993, 478);
			this.mainForm_Fill_Panel.TabIndex = 0;
			// 
			// viewPanel
			// 
			this.viewPanel.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
			this.viewPanel.Controls.Add(this.pageConfiguration);
			this.viewPanel.Controls.Add(this.pageSettings);
			this.viewPanel.Controls.Add(this.pageLogs);
			this.viewPanel.Controls.Add(this.pageMessages);
			this.viewPanel.Location = new System.Drawing.Point(4, 4);
			this.viewPanel.Name = "viewPanel";
			this.viewPanel.Size = new System.Drawing.Size(985, 470);
			this.viewPanel.TabIndex = 3;
			// 
			// _mainForm_Toolbars_Dock_Area_Left
			// 
			this._mainForm_Toolbars_Dock_Area_Left.AccessibleRole = System.Windows.Forms.AccessibleRole.Grouping;
			this._mainForm_Toolbars_Dock_Area_Left.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(191)))), ((int)(((byte)(219)))), ((int)(((byte)(255)))));
			this._mainForm_Toolbars_Dock_Area_Left.DockedPosition = Infragistics.Win.UltraWinToolbars.DockedPosition.Left;
			this._mainForm_Toolbars_Dock_Area_Left.ForeColor = System.Drawing.SystemColors.ControlText;
			this._mainForm_Toolbars_Dock_Area_Left.Location = new System.Drawing.Point(0, 127);
			this._mainForm_Toolbars_Dock_Area_Left.Name = "_mainForm_Toolbars_Dock_Area_Left";
			this._mainForm_Toolbars_Dock_Area_Left.Size = new System.Drawing.Size(0, 478);
			this._mainForm_Toolbars_Dock_Area_Left.ToolbarsManager = this.toolbarsManager;
			// 
			// _mainForm_Toolbars_Dock_Area_Right
			// 
			this._mainForm_Toolbars_Dock_Area_Right.AccessibleRole = System.Windows.Forms.AccessibleRole.Grouping;
			this._mainForm_Toolbars_Dock_Area_Right.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(191)))), ((int)(((byte)(219)))), ((int)(((byte)(255)))));
			this._mainForm_Toolbars_Dock_Area_Right.DockedPosition = Infragistics.Win.UltraWinToolbars.DockedPosition.Right;
			this._mainForm_Toolbars_Dock_Area_Right.ForeColor = System.Drawing.SystemColors.ControlText;
			this._mainForm_Toolbars_Dock_Area_Right.Location = new System.Drawing.Point(993, 127);
			this._mainForm_Toolbars_Dock_Area_Right.Name = "_mainForm_Toolbars_Dock_Area_Right";
			this._mainForm_Toolbars_Dock_Area_Right.Size = new System.Drawing.Size(0, 478);
			this._mainForm_Toolbars_Dock_Area_Right.ToolbarsManager = this.toolbarsManager;
			// 
			// _mainForm_Toolbars_Dock_Area_Top
			// 
			this._mainForm_Toolbars_Dock_Area_Top.AccessibleRole = System.Windows.Forms.AccessibleRole.Grouping;
			this._mainForm_Toolbars_Dock_Area_Top.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(191)))), ((int)(((byte)(219)))), ((int)(((byte)(255)))));
			this._mainForm_Toolbars_Dock_Area_Top.DockedPosition = Infragistics.Win.UltraWinToolbars.DockedPosition.Top;
			this._mainForm_Toolbars_Dock_Area_Top.ForeColor = System.Drawing.SystemColors.ControlText;
			this._mainForm_Toolbars_Dock_Area_Top.Location = new System.Drawing.Point(0, 0);
			this._mainForm_Toolbars_Dock_Area_Top.Name = "_mainForm_Toolbars_Dock_Area_Top";
			this._mainForm_Toolbars_Dock_Area_Top.Size = new System.Drawing.Size(993, 127);
			this._mainForm_Toolbars_Dock_Area_Top.ToolbarsManager = this.toolbarsManager;
			// 
			// _mainForm_Toolbars_Dock_Area_Bottom
			// 
			this._mainForm_Toolbars_Dock_Area_Bottom.AccessibleRole = System.Windows.Forms.AccessibleRole.Grouping;
			this._mainForm_Toolbars_Dock_Area_Bottom.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(191)))), ((int)(((byte)(219)))), ((int)(((byte)(255)))));
			this._mainForm_Toolbars_Dock_Area_Bottom.DockedPosition = Infragistics.Win.UltraWinToolbars.DockedPosition.Bottom;
			this._mainForm_Toolbars_Dock_Area_Bottom.ForeColor = System.Drawing.SystemColors.ControlText;
			this._mainForm_Toolbars_Dock_Area_Bottom.Location = new System.Drawing.Point(0, 605);
			this._mainForm_Toolbars_Dock_Area_Bottom.Name = "_mainForm_Toolbars_Dock_Area_Bottom";
			this._mainForm_Toolbars_Dock_Area_Bottom.Size = new System.Drawing.Size(993, 0);
			this._mainForm_Toolbars_Dock_Area_Bottom.ToolbarsManager = this.toolbarsManager;
			// 
			// pageConfiguration
			// 
			this.pageConfiguration.Location = new System.Drawing.Point(437, 62);
			this.pageConfiguration.Name = "pageConfiguration";
			this.pageConfiguration.Size = new System.Drawing.Size(684, 386);
			this.pageConfiguration.TabIndex = 4;
			// 
			// pageSettings
			// 
			this.pageSettings.Location = new System.Drawing.Point(115, 189);
			this.pageSettings.Name = "pageSettings";
			this.pageSettings.Size = new System.Drawing.Size(505, 130);
			this.pageSettings.TabIndex = 3;
			// 
			// pageLogs
			// 
			this.pageLogs.Location = new System.Drawing.Point(13, 169);
			this.pageLogs.Name = "pageLogs";
			this.pageLogs.Size = new System.Drawing.Size(312, 133);
			this.pageLogs.TabIndex = 2;
			// 
			// pageMessages
			// 
			this.pageMessages.Location = new System.Drawing.Point(13, 12);
			this.pageMessages.Name = "pageMessages";
			this.pageMessages.Size = new System.Drawing.Size(312, 151);
			this.pageMessages.TabIndex = 0;
			// 
			// toolbarsManager
			// 
			this.toolbarsManager.DesignerFlags = 1;
			this.toolbarsManager.DockWithinContainer = this;
			this.toolbarsManager.DockWithinContainerBaseType = typeof(System.Windows.Forms.Form);
			this.toolbarsManager.FormDisplayStyle = Infragistics.Win.UltraWinToolbars.FormDisplayStyle.Standard;
			this.toolbarsManager.Office2007UICompatibility = false;
			optionSet1.AllowAllUp = false;
			this.toolbarsManager.OptionSets.Add(optionSet1);
			this.toolbarsManager.Ribbon.FileMenuStyle = Infragistics.Win.UltraWinToolbars.FileMenuStyle.None;
			ribbonTab1.Caption = "Config";
			ribbonGroup1.Caption = "File";
			buttonTool1.InstanceProps.PreferredSizeOnRibbon = Infragistics.Win.UltraWinToolbars.RibbonToolSize.Large;
			ribbonGroup1.Tools.AddRange(new Infragistics.Win.UltraWinToolbars.ToolBase[] {
            buttonTool1});
			ribbonGroup2.Caption = "Show";
			stateButtonTool1.Checked = true;
			stateButtonTool1.InstanceProps.PreferredSizeOnRibbon = Infragistics.Win.UltraWinToolbars.RibbonToolSize.Large;
			stateButtonTool2.InstanceProps.PreferredSizeOnRibbon = Infragistics.Win.UltraWinToolbars.RibbonToolSize.Large;
			stateButtonTool3.InstanceProps.PreferredSizeOnRibbon = Infragistics.Win.UltraWinToolbars.RibbonToolSize.Large;
			ribbonGroup2.Tools.AddRange(new Infragistics.Win.UltraWinToolbars.ToolBase[] {
            stateButtonTool1,
            stateButtonTool2,
            stateButtonTool3});
			ribbonGroup3.Caption = "Log";
			ribbonGroup3.Tools.AddRange(new Infragistics.Win.UltraWinToolbars.ToolBase[] {
            buttonTool3,
            buttonTool5});
			ribbonTab1.Groups.AddRange(new Infragistics.Win.UltraWinToolbars.RibbonGroup[] {
            ribbonGroup1,
            ribbonGroup2,
            ribbonGroup3});
			ribbonTab2.Caption = "System Settings";
			ribbonGroup4.Caption = "File";
			buttonTool14.InstanceProps.PreferredSizeOnRibbon = Infragistics.Win.UltraWinToolbars.RibbonToolSize.Large;
			ribbonGroup4.Tools.AddRange(new Infragistics.Win.UltraWinToolbars.ToolBase[] {
            buttonTool14});
			ribbonGroup5.Caption = "System";
			buttonTool9.InstanceProps.PreferredSizeOnRibbon = Infragistics.Win.UltraWinToolbars.RibbonToolSize.Large;
			buttonTool13.InstanceProps.PreferredSizeOnRibbon = Infragistics.Win.UltraWinToolbars.RibbonToolSize.Large;
			ribbonGroup5.Tools.AddRange(new Infragistics.Win.UltraWinToolbars.ToolBase[] {
            buttonTool9,
            buttonTool13});
			ribbonTab2.Groups.AddRange(new Infragistics.Win.UltraWinToolbars.RibbonGroup[] {
            ribbonGroup4,
            ribbonGroup5});
			this.toolbarsManager.Ribbon.NonInheritedRibbonTabs.AddRange(new Infragistics.Win.UltraWinToolbars.RibbonTab[] {
            ribbonTab1,
            ribbonTab2});
			this.toolbarsManager.Ribbon.QuickAccessToolbar.Visible = false;
			this.toolbarsManager.Ribbon.Visible = true;
			this.toolbarsManager.ShowFullMenusDelay = 500;
			this.toolbarsManager.Style = Infragistics.Win.UltraWinToolbars.ToolbarStyle.Office2007;
			appearance1.Image = global::Fusion.Properties.Resources.Save_32;
			buttonTool2.SharedPropsInternal.AppearancesLarge.Appearance = appearance1;
			buttonTool2.SharedPropsInternal.Caption = "Save";
			buttonTool2.SharedPropsInternal.Category = "File";
			stateButtonTool4.Checked = true;
			stateButtonTool4.OptionSetKey = "Show";
			appearance2.Image = global::Fusion.Properties.Resources.Configuration_32;
			stateButtonTool4.SharedPropsInternal.AppearancesLarge.Appearance = appearance2;
			stateButtonTool4.SharedPropsInternal.Caption = "Configuration";
			stateButtonTool4.SharedPropsInternal.Category = "Show";
			stateButtonTool5.OptionSetKey = "Show";
			appearance3.Image = global::Fusion.Properties.Resources.Message_32;
			stateButtonTool5.SharedPropsInternal.AppearancesLarge.Appearance = appearance3;
			stateButtonTool5.SharedPropsInternal.Caption = "Messages";
			stateButtonTool5.SharedPropsInternal.Category = "Show";
			stateButtonTool6.OptionSetKey = "Show";
			appearance4.Image = global::Fusion.Properties.Resources.Log_32;
			stateButtonTool6.SharedPropsInternal.AppearancesLarge.Appearance = appearance4;
			stateButtonTool6.SharedPropsInternal.Caption = "Logs";
			stateButtonTool6.SharedPropsInternal.Category = "Show";
			buttonTool4.SharedPropsInternal.Caption = "Purge";
			buttonTool4.SharedPropsInternal.Category = "Logs";
			buttonTool6.SharedPropsInternal.Caption = "Export";
			buttonTool6.SharedPropsInternal.Category = "Logs";
			appearance5.Image = global::Fusion.Properties.Resources.Start_32;
			buttonTool11.SharedPropsInternal.AppearancesLarge.Appearance = appearance5;
			buttonTool11.SharedPropsInternal.Caption = "Start Connector";
			appearance6.Image = global::Fusion.Properties.Resources.Stop_32;
			buttonTool12.SharedPropsInternal.AppearancesLarge.Appearance = appearance6;
			buttonTool12.SharedPropsInternal.Caption = "Stop Connector";
			this.toolbarsManager.Tools.AddRange(new Infragistics.Win.UltraWinToolbars.ToolBase[] {
            buttonTool2,
            stateButtonTool4,
            stateButtonTool5,
            stateButtonTool6,
            buttonTool4,
            buttonTool6,
            buttonTool11,
            buttonTool12});
			this.toolbarsManager.BeforeRibbonTabSelected += new Infragistics.Win.UltraWinToolbars.BeforeRibbonTabSelectedEventHandler(this.ToolbarsManagerBeforeRibbonTabSelected);
			this.toolbarsManager.ToolClick += new Infragistics.Win.UltraWinToolbars.ToolClickEventHandler(this.ToolbarsManagerToolClick);
			// 
			// MainForm
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.ClientSize = new System.Drawing.Size(993, 605);
			this.Controls.Add(this.mainForm_Fill_Panel);
			this.Controls.Add(this._mainForm_Toolbars_Dock_Area_Left);
			this.Controls.Add(this._mainForm_Toolbars_Dock_Area_Right);
			this.Controls.Add(this._mainForm_Toolbars_Dock_Area_Bottom);
			this.Controls.Add(this._mainForm_Toolbars_Dock_Area_Top);
			this.Name = "MainForm";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "TITLE";
			this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.MainFormFormClosing);
			this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.MainFormFormClosed);
			this.Load += new System.EventHandler(this.FrmMainLoad);
			this.mainForm_Fill_Panel.ClientArea.ResumeLayout(false);
			this.mainForm_Fill_Panel.ResumeLayout(false);
			this.viewPanel.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.toolbarsManager)).EndInit();
			this.ResumeLayout(false);

        }

        #endregion

        private Infragistics.Win.UltraWinToolbars.UltraToolbarsManager toolbarsManager;
        private Infragistics.Win.Misc.UltraPanel mainForm_Fill_Panel;
        private Infragistics.Win.UltraWinToolbars.UltraToolbarsDockArea _mainForm_Toolbars_Dock_Area_Left;
        private Infragistics.Win.UltraWinToolbars.UltraToolbarsDockArea _mainForm_Toolbars_Dock_Area_Right;
        private Infragistics.Win.UltraWinToolbars.UltraToolbarsDockArea _mainForm_Toolbars_Dock_Area_Bottom;
		  private Infragistics.Win.UltraWinToolbars.UltraToolbarsDockArea _mainForm_Toolbars_Dock_Area_Top;
		  private Pages.Logs pageLogs;
        private Pages.Messages pageMessages;
		  private System.Windows.Forms.Panel viewPanel;
		  private Pages.Settings pageSettings;
		  private Pages.Configuration pageConfiguration;
		  private System.Windows.Forms.Timer timer1;
    }
}