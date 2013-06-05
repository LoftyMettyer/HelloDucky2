namespace Fusion.Pages
{
    partial class Messages
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
			Infragistics.Win.UltraWinTree.UltraTreeColumnSet ultraTreeColumnSet1 = new Infragistics.Win.UltraWinTree.UltraTreeColumnSet();
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinTree.UltraTreeNodeColumn ultraTreeNodeColumn1 = new Infragistics.Win.UltraWinTree.UltraTreeNodeColumn();
			Infragistics.Win.UltraWinTree.UltraTreeNodeColumn ultraTreeNodeColumn2 = new Infragistics.Win.UltraWinTree.UltraTreeNodeColumn();
			Infragistics.Win.UltraWinTree.UltraTreeNodeColumn ultraTreeNodeColumn3 = new Infragistics.Win.UltraWinTree.UltraTreeNodeColumn();
			Infragistics.Win.UltraWinTree.UltraTreeNodeColumn ultraTreeNodeColumn4 = new Infragistics.Win.UltraWinTree.UltraTreeNodeColumn();
			Infragistics.Win.UltraWinTree.UltraTreeNodeColumn ultraTreeNodeColumn5 = new Infragistics.Win.UltraWinTree.UltraTreeNodeColumn();
			Infragistics.Win.UltraWinTree.UltraTreeColumnSet ultraTreeColumnSet2 = new Infragistics.Win.UltraWinTree.UltraTreeColumnSet();
			Infragistics.Win.UltraWinTree.UltraTreeNodeColumn ultraTreeNodeColumn6 = new Infragistics.Win.UltraWinTree.UltraTreeNodeColumn();
			Infragistics.Win.UltraWinTree.UltraTreeNodeColumn ultraTreeNodeColumn7 = new Infragistics.Win.UltraWinTree.UltraTreeNodeColumn();
			this.messageTree = new Infragistics.Win.UltraWinTree.UltraTree();
			this.messageBindingSource = new System.Windows.Forms.BindingSource(this.components);
			this.fieldPanel = new System.Windows.Forms.Panel();
			this.SubscribeEditor = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.PublishEditor = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.TemplateXmlEditor = new Infragistics.Win.FormattedLinkLabel.UltraFormattedTextEditor();
			this.splitter1 = new System.Windows.Forms.Splitter();
			((System.ComponentModel.ISupportInitialize)(this.messageTree)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.messageBindingSource)).BeginInit();
			this.fieldPanel.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.SubscribeEditor)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.PublishEditor)).BeginInit();
			this.SuspendLayout();
			// 
			// messageTree
			// 
			appearance1.ImageVAlign = Infragistics.Win.VAlign.Bottom;
			ultraTreeColumnSet1.CellAppearance = appearance1;
			ultraTreeNodeColumn1.ButtonDisplayStyle = Infragistics.Win.UltraWinTree.ButtonDisplayStyle.Always;
			ultraTreeNodeColumn1.DataType = typeof(string);
			ultraTreeNodeColumn1.Key = "Name";
			ultraTreeNodeColumn1.LayoutInfo.PreferredCellSize = new System.Drawing.Size(93, 16);
			ultraTreeNodeColumn1.LayoutInfo.PreferredLabelSize = new System.Drawing.Size(93, 0);
			ultraTreeNodeColumn1.SortType = Infragistics.Win.UltraWinTree.SortType.Ascending;
			ultraTreeNodeColumn2.ButtonDisplayStyle = Infragistics.Win.UltraWinTree.ButtonDisplayStyle.Always;
			ultraTreeNodeColumn2.DataType = typeof(bool);
			ultraTreeNodeColumn2.Key = "AllowPublish";
			ultraTreeNodeColumn3.ButtonDisplayStyle = Infragistics.Win.UltraWinTree.ButtonDisplayStyle.Always;
			ultraTreeNodeColumn3.DataType = typeof(bool);
			ultraTreeNodeColumn3.Key = "AllowSubscribe";
			ultraTreeNodeColumn4.ButtonDisplayStyle = Infragistics.Win.UltraWinTree.ButtonDisplayStyle.Always;
			ultraTreeNodeColumn4.DataType = typeof(string);
			ultraTreeNodeColumn4.Key = "XmlTemplate";
			ultraTreeNodeColumn5.ButtonDisplayStyle = Infragistics.Win.UltraWinTree.ButtonDisplayStyle.Always;
			ultraTreeNodeColumn5.DataType = typeof(System.ComponentModel.BindingList<Fusion.Pages.FusionMessageElementNode>);
			ultraTreeNodeColumn5.IsChaptered = true;
			ultraTreeNodeColumn5.Key = "Items";
			ultraTreeColumnSet1.Columns.Add(ultraTreeNodeColumn1);
			ultraTreeColumnSet1.Columns.Add(ultraTreeNodeColumn2);
			ultraTreeColumnSet1.Columns.Add(ultraTreeNodeColumn3);
			ultraTreeColumnSet1.Columns.Add(ultraTreeNodeColumn4);
			ultraTreeColumnSet1.Columns.Add(ultraTreeNodeColumn5);
			ultraTreeColumnSet1.IsAutoGenerated = true;
			ultraTreeNodeColumn6.ButtonDisplayStyle = Infragistics.Win.UltraWinTree.ButtonDisplayStyle.Always;
			ultraTreeNodeColumn6.DataType = typeof(string);
			ultraTreeNodeColumn6.Key = "Name";
			ultraTreeNodeColumn7.ButtonDisplayStyle = Infragistics.Win.UltraWinTree.ButtonDisplayStyle.Always;
			ultraTreeNodeColumn7.DataType = typeof(int);
			ultraTreeNodeColumn7.Key = "Position";
			ultraTreeNodeColumn7.SortType = Infragistics.Win.UltraWinTree.SortType.Ascending;
			ultraTreeColumnSet2.Columns.Add(ultraTreeNodeColumn6);
			ultraTreeColumnSet2.Columns.Add(ultraTreeNodeColumn7);
			ultraTreeColumnSet2.IsAutoGenerated = true;
			ultraTreeColumnSet2.Key = "Items";
			this.messageTree.ColumnSettings.ColumnSets.Add(ultraTreeColumnSet1);
			this.messageTree.ColumnSettings.ColumnSets.Add(ultraTreeColumnSet2);
			this.messageTree.DataSource = this.messageBindingSource;
			this.messageTree.Dock = System.Windows.Forms.DockStyle.Left;
			this.messageTree.DrawsFocusRect = Infragistics.Win.DefaultableBoolean.True;
			this.messageTree.FullRowSelect = true;
			this.messageTree.Location = new System.Drawing.Point(0, 0);
			this.messageTree.Name = "messageTree";
			this.messageTree.Size = new System.Drawing.Size(262, 427);
			this.messageTree.SynchronizeCurrencyManager = true;
			this.messageTree.TabIndex = 0;
			this.messageTree.ViewStyle = Infragistics.Win.UltraWinTree.ViewStyle.Grid;
			// 
			// messageBindingSource
			// 
			this.messageBindingSource.DataSource = typeof(Fusion.Pages.FusionMessageNode);
			// 
			// fieldPanel
			// 
			this.fieldPanel.Controls.Add(this.SubscribeEditor);
			this.fieldPanel.Controls.Add(this.PublishEditor);
			this.fieldPanel.Dock = System.Windows.Forms.DockStyle.Top;
			this.fieldPanel.Location = new System.Drawing.Point(266, 0);
			this.fieldPanel.Name = "fieldPanel";
			this.fieldPanel.Size = new System.Drawing.Size(482, 44);
			this.fieldPanel.TabIndex = 1;
			// 
			// SubscribeEditor
			// 
			this.SubscribeEditor.DataBindings.Add(new System.Windows.Forms.Binding("Checked", this.messageBindingSource, "Subscribe", true));
			this.SubscribeEditor.DataBindings.Add(new System.Windows.Forms.Binding("Enabled", this.messageBindingSource, "AllowSubscribe", true, System.Windows.Forms.DataSourceUpdateMode.Never));
			this.SubscribeEditor.Location = new System.Drawing.Point(89, 11);
			this.SubscribeEditor.Name = "SubscribeEditor";
			this.SubscribeEditor.Size = new System.Drawing.Size(81, 20);
			this.SubscribeEditor.TabIndex = 1;
			this.SubscribeEditor.Text = "Subscribe";
			// 
			// PublishEditor
			// 
			this.PublishEditor.DataBindings.Add(new System.Windows.Forms.Binding("Checked", this.messageBindingSource, "Publish", true));
			this.PublishEditor.DataBindings.Add(new System.Windows.Forms.Binding("Enabled", this.messageBindingSource, "AllowPublish", true, System.Windows.Forms.DataSourceUpdateMode.Never));
			this.PublishEditor.Location = new System.Drawing.Point(12, 11);
			this.PublishEditor.Name = "PublishEditor";
			this.PublishEditor.Size = new System.Drawing.Size(71, 20);
			this.PublishEditor.TabIndex = 0;
			this.PublishEditor.Text = "Publish";
			// 
			// TemplateXmlEditor
			// 
			this.TemplateXmlEditor.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.messageBindingSource, "XmlTemplate", true));
			this.TemplateXmlEditor.Dock = System.Windows.Forms.DockStyle.Fill;
			this.TemplateXmlEditor.Location = new System.Drawing.Point(266, 44);
			this.TemplateXmlEditor.Name = "TemplateXmlEditor";
			this.TemplateXmlEditor.ReadOnly = true;
			this.TemplateXmlEditor.Size = new System.Drawing.Size(482, 383);
			this.TemplateXmlEditor.TabIndex = 2;
			this.TemplateXmlEditor.Value = "";
			// 
			// splitter1
			// 
			this.splitter1.Location = new System.Drawing.Point(262, 0);
			this.splitter1.Name = "splitter1";
			this.splitter1.Size = new System.Drawing.Size(4, 427);
			this.splitter1.TabIndex = 6;
			this.splitter1.TabStop = false;
			// 
			// Messages
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.Controls.Add(this.TemplateXmlEditor);
			this.Controls.Add(this.fieldPanel);
			this.Controls.Add(this.splitter1);
			this.Controls.Add(this.messageTree);
			this.Name = "Messages";
			this.Size = new System.Drawing.Size(748, 427);
			((System.ComponentModel.ISupportInitialize)(this.messageTree)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.messageBindingSource)).EndInit();
			this.fieldPanel.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.SubscribeEditor)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.PublishEditor)).EndInit();
			this.ResumeLayout(false);

        }

        #endregion

		  private Infragistics.Win.UltraWinTree.UltraTree messageTree;
		  private System.Windows.Forms.BindingSource messageBindingSource;
		  private System.Windows.Forms.Panel fieldPanel;
		  private Infragistics.Win.UltraWinEditors.UltraCheckEditor SubscribeEditor;
		  private Infragistics.Win.UltraWinEditors.UltraCheckEditor PublishEditor;
		  private Infragistics.Win.FormattedLinkLabel.UltraFormattedTextEditor TemplateXmlEditor;
		  private System.Windows.Forms.Splitter splitter1;
    }
}
