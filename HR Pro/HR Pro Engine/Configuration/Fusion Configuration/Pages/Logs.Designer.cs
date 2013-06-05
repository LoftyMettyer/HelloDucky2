namespace Fusion.Pages
{
    partial class Logs
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

        #region Component Designer generated code

        /// <summary> 
        /// Required method for Designer support - do not modify 
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
			this.components = new System.ComponentModel.Container();
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("FusionLog", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("MessageType");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("BusRef");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("LastGeneratedDate", -1, null, 0, Infragistics.Win.UltraWinGrid.SortIndicator.Descending, false);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("LastProcessedDate");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("LastGeneratedXml");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("UserName");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Id");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			this.LogGrid = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.logBindingSource = new System.Windows.Forms.BindingSource(this.components);
			this.panel1 = new System.Windows.Forms.Panel();
			this.FindButton = new Infragistics.Win.Misc.UltraButton();
			this.ultraLabel3 = new Infragistics.Win.Misc.UltraLabel();
			this.ultraLabel2 = new Infragistics.Win.Misc.UltraLabel();
			this.messageTypeEditor = new Infragistics.Win.UltraWinEditors.UltraComboEditor();
			this.queryBindingSource = new System.Windows.Forms.BindingSource(this.components);
			this.DateLastGeneratedToEditor = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.DateLastGeneratedFromEditor = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.ultraLabel1 = new Infragistics.Win.Misc.UltraLabel();
			this.splitter1 = new System.Windows.Forms.Splitter();
			this.LastGeneratedXmlEditor = new Infragistics.Win.FormattedLinkLabel.UltraFormattedTextEditor();
			((System.ComponentModel.ISupportInitialize)(this.LogGrid)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.logBindingSource)).BeginInit();
			this.panel1.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.messageTypeEditor)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.queryBindingSource)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.DateLastGeneratedToEditor)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.DateLastGeneratedFromEditor)).BeginInit();
			this.SuspendLayout();
			// 
			// LogGrid
			// 
			this.LogGrid.DataSource = this.logBindingSource;
			appearance1.BackColor = System.Drawing.SystemColors.Window;
			appearance1.BorderColor = System.Drawing.SystemColors.InactiveCaption;
			this.LogGrid.DisplayLayout.Appearance = appearance1;
			ultraGridColumn2.AutoCompleteMode = Infragistics.Win.AutoCompleteMode.Append;
			ultraGridColumn2.Header.Caption = "Message Type";
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 173;
			ultraGridColumn3.AutoCompleteMode = Infragistics.Win.AutoCompleteMode.Append;
			ultraGridColumn3.Header.Caption = "Bus Reference";
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Width = 225;
			ultraGridColumn1.AutoCompleteMode = Infragistics.Win.AutoCompleteMode.Append;
			ultraGridColumn1.Header.Caption = "Last Generated";
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Width = 99;
			ultraGridColumn4.AutoCompleteMode = Infragistics.Win.AutoCompleteMode.Append;
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn5.AutoCompleteMode = Infragistics.Win.AutoCompleteMode.Append;
			ultraGridColumn5.Header.VisiblePosition = 4;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn6.AutoCompleteMode = Infragistics.Win.AutoCompleteMode.Append;
			ultraGridColumn6.Header.VisiblePosition = 5;
			ultraGridColumn6.Hidden = true;
			ultraGridColumn7.AutoCompleteMode = Infragistics.Win.AutoCompleteMode.Append;
			ultraGridColumn7.Header.VisiblePosition = 6;
			ultraGridColumn7.Hidden = true;
			ultraGridBand1.Columns.AddRange(new object[] {
            ultraGridColumn2,
            ultraGridColumn3,
            ultraGridColumn1,
            ultraGridColumn4,
            ultraGridColumn5,
            ultraGridColumn6,
            ultraGridColumn7});
			this.LogGrid.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.LogGrid.DisplayLayout.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.LogGrid.DisplayLayout.CaptionVisible = Infragistics.Win.DefaultableBoolean.False;
			appearance2.BackColor = System.Drawing.SystemColors.ActiveBorder;
			appearance2.BackColor2 = System.Drawing.SystemColors.ControlDark;
			appearance2.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance2.BorderColor = System.Drawing.SystemColors.Window;
			this.LogGrid.DisplayLayout.GroupByBox.Appearance = appearance2;
			appearance3.ForeColor = System.Drawing.SystemColors.GrayText;
			this.LogGrid.DisplayLayout.GroupByBox.BandLabelAppearance = appearance3;
			this.LogGrid.DisplayLayout.GroupByBox.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			appearance4.BackColor = System.Drawing.SystemColors.ControlLightLight;
			appearance4.BackColor2 = System.Drawing.SystemColors.Control;
			appearance4.BackGradientStyle = Infragistics.Win.GradientStyle.Horizontal;
			appearance4.ForeColor = System.Drawing.SystemColors.GrayText;
			this.LogGrid.DisplayLayout.GroupByBox.PromptAppearance = appearance4;
			this.LogGrid.DisplayLayout.MaxColScrollRegions = 1;
			this.LogGrid.DisplayLayout.MaxRowScrollRegions = 1;
			appearance5.BackColor = System.Drawing.SystemColors.Window;
			appearance5.ForeColor = System.Drawing.SystemColors.ControlText;
			this.LogGrid.DisplayLayout.Override.ActiveCellAppearance = appearance5;
			appearance6.BackColor = System.Drawing.SystemColors.Highlight;
			appearance6.ForeColor = System.Drawing.SystemColors.HighlightText;
			this.LogGrid.DisplayLayout.Override.ActiveRowAppearance = appearance6;
			this.LogGrid.DisplayLayout.Override.BorderStyleCell = Infragistics.Win.UIElementBorderStyle.Dotted;
			this.LogGrid.DisplayLayout.Override.BorderStyleRow = Infragistics.Win.UIElementBorderStyle.Dotted;
			appearance7.BackColor = System.Drawing.SystemColors.Window;
			this.LogGrid.DisplayLayout.Override.CardAreaAppearance = appearance7;
			appearance8.BorderColor = System.Drawing.Color.Silver;
			appearance8.TextTrimming = Infragistics.Win.TextTrimming.EllipsisCharacter;
			this.LogGrid.DisplayLayout.Override.CellAppearance = appearance8;
			this.LogGrid.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.EditAndSelectText;
			this.LogGrid.DisplayLayout.Override.CellPadding = 0;
			appearance9.BackColor = System.Drawing.SystemColors.Control;
			appearance9.BackColor2 = System.Drawing.SystemColors.ControlDark;
			appearance9.BackGradientAlignment = Infragistics.Win.GradientAlignment.Element;
			appearance9.BackGradientStyle = Infragistics.Win.GradientStyle.Horizontal;
			appearance9.BorderColor = System.Drawing.SystemColors.Window;
			this.LogGrid.DisplayLayout.Override.GroupByRowAppearance = appearance9;
			appearance10.TextHAlignAsString = "Left";
			this.LogGrid.DisplayLayout.Override.HeaderAppearance = appearance10;
			this.LogGrid.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			this.LogGrid.DisplayLayout.Override.HeaderStyle = Infragistics.Win.HeaderStyle.WindowsXPCommand;
			appearance11.BackColor = System.Drawing.SystemColors.Window;
			appearance11.BorderColor = System.Drawing.Color.Silver;
			this.LogGrid.DisplayLayout.Override.RowAppearance = appearance11;
			this.LogGrid.DisplayLayout.Override.RowSelectors = Infragistics.Win.DefaultableBoolean.False;
			this.LogGrid.DisplayLayout.Override.SelectTypeRow = Infragistics.Win.UltraWinGrid.SelectType.Single;
			appearance12.BackColor = System.Drawing.SystemColors.ControlLight;
			this.LogGrid.DisplayLayout.Override.TemplateAddRowAppearance = appearance12;
			this.LogGrid.DisplayLayout.ScrollBounds = Infragistics.Win.UltraWinGrid.ScrollBounds.ScrollToFill;
			this.LogGrid.DisplayLayout.ScrollStyle = Infragistics.Win.UltraWinGrid.ScrollStyle.Immediate;
			this.LogGrid.DisplayLayout.ViewStyle = Infragistics.Win.UltraWinGrid.ViewStyle.SingleBand;
			this.LogGrid.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
			this.LogGrid.Dock = System.Windows.Forms.DockStyle.Left;
			this.LogGrid.Location = new System.Drawing.Point(0, 63);
			this.LogGrid.Name = "LogGrid";
			this.LogGrid.Size = new System.Drawing.Size(527, 556);
			this.LogGrid.TabIndex = 1;
			this.LogGrid.Text = "ultraGrid1";
			// 
			// logBindingSource
			// 
			this.logBindingSource.DataSource = typeof(Fusion.FusionLog);
			// 
			// panel1
			// 
			this.panel1.Controls.Add(this.FindButton);
			this.panel1.Controls.Add(this.ultraLabel3);
			this.panel1.Controls.Add(this.ultraLabel2);
			this.panel1.Controls.Add(this.messageTypeEditor);
			this.panel1.Controls.Add(this.DateLastGeneratedToEditor);
			this.panel1.Controls.Add(this.DateLastGeneratedFromEditor);
			this.panel1.Controls.Add(this.ultraLabel1);
			this.panel1.Dock = System.Windows.Forms.DockStyle.Top;
			this.panel1.Location = new System.Drawing.Point(0, 0);
			this.panel1.Name = "panel1";
			this.panel1.Size = new System.Drawing.Size(928, 63);
			this.panel1.TabIndex = 0;
			// 
			// FindButton
			// 
			this.FindButton.Location = new System.Drawing.Point(318, 4);
			this.FindButton.Name = "FindButton";
			this.FindButton.Size = new System.Drawing.Size(55, 23);
			this.FindButton.TabIndex = 6;
			this.FindButton.Text = "Find";
			// 
			// ultraLabel3
			// 
			this.ultraLabel3.AutoSize = true;
			this.ultraLabel3.Location = new System.Drawing.Point(193, 34);
			this.ultraLabel3.Name = "ultraLabel3";
			this.ultraLabel3.Size = new System.Drawing.Size(14, 14);
			this.ultraLabel3.TabIndex = 4;
			this.ultraLabel3.Text = "to";
			// 
			// ultraLabel2
			// 
			this.ultraLabel2.AutoSize = true;
			this.ultraLabel2.Location = new System.Drawing.Point(5, 34);
			this.ultraLabel2.Name = "ultraLabel2";
			this.ultraLabel2.Size = new System.Drawing.Size(85, 14);
			this.ultraLabel2.TabIndex = 2;
			this.ultraLabel2.Text = "Last Generated:";
			// 
			// messageTypeEditor
			// 
			this.messageTypeEditor.DataBindings.Add(new System.Windows.Forms.Binding("Text", this.queryBindingSource, "MessageType", true, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged));
			this.messageTypeEditor.DropDownStyle = Infragistics.Win.DropDownStyle.DropDownList;
			this.messageTypeEditor.Location = new System.Drawing.Point(96, 4);
			this.messageTypeEditor.Name = "messageTypeEditor";
			this.messageTypeEditor.Size = new System.Drawing.Size(207, 21);
			this.messageTypeEditor.TabIndex = 1;
			// 
			// queryBindingSource
			// 
			this.queryBindingSource.DataSource = typeof(Fusion.Pages.FusionLogQuery);
			// 
			// DateLastGeneratedToEditor
			// 
			this.DateLastGeneratedToEditor.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.queryBindingSource, "DateLastGeneratedTo", true, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged));
			this.DateLastGeneratedToEditor.DateButtons.Add(dateButton1);
			this.DateLastGeneratedToEditor.Location = new System.Drawing.Point(212, 31);
			this.DateLastGeneratedToEditor.Name = "DateLastGeneratedToEditor";
			this.DateLastGeneratedToEditor.NonAutoSizeHeight = 21;
			this.DateLastGeneratedToEditor.NullDateLabel = "";
			this.DateLastGeneratedToEditor.Size = new System.Drawing.Size(91, 21);
			this.DateLastGeneratedToEditor.TabIndex = 5;
			this.DateLastGeneratedToEditor.Value = "";
			// 
			// DateLastGeneratedFromEditor
			// 
			this.DateLastGeneratedFromEditor.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.queryBindingSource, "DateLastGeneratedFrom", true, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged));
			this.DateLastGeneratedFromEditor.DateButtons.Add(dateButton2);
			this.DateLastGeneratedFromEditor.Location = new System.Drawing.Point(96, 31);
			this.DateLastGeneratedFromEditor.Name = "DateLastGeneratedFromEditor";
			this.DateLastGeneratedFromEditor.NonAutoSizeHeight = 21;
			this.DateLastGeneratedFromEditor.NullDateLabel = "";
			this.DateLastGeneratedFromEditor.Size = new System.Drawing.Size(91, 21);
			this.DateLastGeneratedFromEditor.TabIndex = 3;
			this.DateLastGeneratedFromEditor.Value = "";
			// 
			// ultraLabel1
			// 
			this.ultraLabel1.AutoSize = true;
			this.ultraLabel1.Location = new System.Drawing.Point(5, 8);
			this.ultraLabel1.Name = "ultraLabel1";
			this.ultraLabel1.Size = new System.Drawing.Size(81, 14);
			this.ultraLabel1.TabIndex = 0;
			this.ultraLabel1.Text = "Message Type:";
			// 
			// splitter1
			// 
			this.splitter1.Location = new System.Drawing.Point(527, 63);
			this.splitter1.Name = "splitter1";
			this.splitter1.Size = new System.Drawing.Size(4, 556);
			this.splitter1.TabIndex = 9;
			this.splitter1.TabStop = false;
			// 
			// LastGeneratedXmlEditor
			// 
			this.LastGeneratedXmlEditor.DataBindings.Add(new System.Windows.Forms.Binding("Text", this.logBindingSource, "LastGeneratedXml", true));
			this.LastGeneratedXmlEditor.Dock = System.Windows.Forms.DockStyle.Fill;
			this.LastGeneratedXmlEditor.Location = new System.Drawing.Point(531, 63);
			this.LastGeneratedXmlEditor.Name = "LastGeneratedXmlEditor";
			this.LastGeneratedXmlEditor.ReadOnly = true;
			this.LastGeneratedXmlEditor.Size = new System.Drawing.Size(397, 556);
			this.LastGeneratedXmlEditor.TabIndex = 10;
			this.LastGeneratedXmlEditor.Value = "";
			// 
			// Logs
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.Controls.Add(this.LastGeneratedXmlEditor);
			this.Controls.Add(this.splitter1);
			this.Controls.Add(this.LogGrid);
			this.Controls.Add(this.panel1);
			this.Name = "Logs";
			this.Size = new System.Drawing.Size(928, 619);
			((System.ComponentModel.ISupportInitialize)(this.LogGrid)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.logBindingSource)).EndInit();
			this.panel1.ResumeLayout(false);
			this.panel1.PerformLayout();
			((System.ComponentModel.ISupportInitialize)(this.messageTypeEditor)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.queryBindingSource)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.DateLastGeneratedToEditor)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.DateLastGeneratedFromEditor)).EndInit();
			this.ResumeLayout(false);

        }

        #endregion

		  private Infragistics.Win.UltraWinGrid.UltraGrid LogGrid;
		  private System.Windows.Forms.Panel panel1;
		  private Infragistics.Win.UltraWinEditors.UltraComboEditor messageTypeEditor;
		  private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo DateLastGeneratedToEditor;
		  private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo DateLastGeneratedFromEditor;
		  private Infragistics.Win.Misc.UltraLabel ultraLabel1;
		  private Infragistics.Win.Misc.UltraLabel ultraLabel3;
		  private Infragistics.Win.Misc.UltraLabel ultraLabel2;
		  private System.Windows.Forms.BindingSource logBindingSource;
		  private System.Windows.Forms.BindingSource queryBindingSource;
		  private Infragistics.Win.Misc.UltraButton FindButton;
		  private System.Windows.Forms.Splitter splitter1;
		  private Infragistics.Win.FormattedLinkLabel.UltraFormattedTextEditor LastGeneratedXmlEditor;
    }
}
