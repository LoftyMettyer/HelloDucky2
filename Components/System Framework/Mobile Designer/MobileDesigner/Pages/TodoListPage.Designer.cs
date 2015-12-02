namespace MobileDesigner.Pages
{
    partial class TodoListPage
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
            this.btnRefresh = new MobileDesigner.Controls.IconButton();
            this.btnCancel = new MobileDesigner.Controls.IconButton();
            this.lblInstruction = new System.Windows.Forms.Label();
            this.controlsPanel = new System.Windows.Forms.TableLayoutPanel();
            this.todoLinkButton = new MobileDesigner.Controls.LinkButton();
            this.lblNothingTodo = new System.Windows.Forms.Label();
            this.buttonsPanel = new System.Windows.Forms.TableLayoutPanel();
            this.Footer.SuspendLayout();
            this.Main.SuspendLayout();
            this.controlsPanel.SuspendLayout();
            this.buttonsPanel.SuspendLayout();
            this.SuspendLayout();
            // 
            // Footer
            // 
            this.Footer.Controls.Add(this.buttonsPanel);
            // 
            // Main
            // 
            this.Main.Controls.Add(this.controlsPanel);
            // 
            // btnRefresh
            // 
            this.btnRefresh.Anchor = System.Windows.Forms.AnchorStyles.Top;
            this.btnRefresh.AutoSize = true;
            this.btnRefresh.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.btnRefresh.Caption = "Refresh";
            this.btnRefresh.Image = null;
            this.btnRefresh.Location = new System.Drawing.Point(32, 3);
            this.btnRefresh.Name = "btnRefresh";
            this.btnRefresh.Size = new System.Drawing.Size(36, 50);
            this.btnRefresh.TabIndex = 0;
            // 
            // btnCancel
            // 
            this.btnCancel.Anchor = System.Windows.Forms.AnchorStyles.Top;
            this.btnCancel.AutoSize = true;
            this.btnCancel.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.btnCancel.Caption = "Cancel";
            this.btnCancel.Image = null;
            this.btnCancel.Location = new System.Drawing.Point(134, 3);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(32, 50);
            this.btnCancel.TabIndex = 1;
            // 
            // lblInstruction
            // 
            this.lblInstruction.AutoSize = true;
            this.lblInstruction.BackColor = System.Drawing.Color.Transparent;
            this.controlsPanel.SetColumnSpan(this.lblInstruction, 2);
            this.lblInstruction.Location = new System.Drawing.Point(3, 13);
            this.lblInstruction.Name = "lblInstruction";
            this.lblInstruction.Size = new System.Drawing.Size(150, 13);
            this.lblInstruction.TabIndex = 0;
            this.lblInstruction.Text = "Welcome Caption - items in list";
            this.lblInstruction.UseMnemonic = false;
            // 
            // controlsPanel
            // 
            this.controlsPanel.ColumnCount = 2;
            this.controlsPanel.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle());
            this.controlsPanel.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.controlsPanel.Controls.Add(this.todoLinkButton, 0, 2);
            this.controlsPanel.Controls.Add(this.lblNothingTodo, 0, 0);
            this.controlsPanel.Controls.Add(this.lblInstruction, 0, 1);
            this.controlsPanel.Location = new System.Drawing.Point(12, 24);
            this.controlsPanel.Name = "controlsPanel";
            this.controlsPanel.RowCount = 4;
            this.controlsPanel.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.controlsPanel.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.controlsPanel.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.controlsPanel.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.controlsPanel.Size = new System.Drawing.Size(258, 212);
            this.controlsPanel.TabIndex = 1;
            // 
            // todoLinkButton
            // 
            this.todoLinkButton.AutoSize = true;
            this.todoLinkButton.BackColor = System.Drawing.Color.Transparent;
            this.controlsPanel.SetColumnSpan(this.todoLinkButton, 2);
            this.todoLinkButton.DescriptionFont = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.todoLinkButton.DescriptionForeColor = System.Drawing.SystemColors.ControlText;
            this.todoLinkButton.Location = new System.Drawing.Point(3, 29);
            this.todoLinkButton.Name = "todoLinkButton";
            this.todoLinkButton.Size = new System.Drawing.Size(98, 32);
            this.todoLinkButton.TabIndex = 4;
            this.todoLinkButton.TitleFont = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.todoLinkButton.TitleForeColor = System.Drawing.SystemColors.ControlText;
            // 
            // lblNothingTodo
            // 
            this.lblNothingTodo.AutoSize = true;
            this.lblNothingTodo.BackColor = System.Drawing.Color.Transparent;
            this.controlsPanel.SetColumnSpan(this.lblNothingTodo, 2);
            this.lblNothingTodo.Location = new System.Drawing.Point(3, 0);
            this.lblNothingTodo.Name = "lblNothingTodo";
            this.lblNothingTodo.Size = new System.Drawing.Size(165, 13);
            this.lblNothingTodo.TabIndex = 2;
            this.lblNothingTodo.Text = "Welcome Caption - no items in list";
            this.lblNothingTodo.UseMnemonic = false;
            // 
            // buttonsPanel
            // 
            this.buttonsPanel.ColumnCount = 2;
            this.buttonsPanel.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.buttonsPanel.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.buttonsPanel.Controls.Add(this.btnRefresh, 0, 0);
            this.buttonsPanel.Controls.Add(this.btnCancel, 1, 0);
            this.buttonsPanel.Location = new System.Drawing.Point(4, 4);
            this.buttonsPanel.Name = "buttonsPanel";
            this.buttonsPanel.RowCount = 1;
            this.buttonsPanel.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.buttonsPanel.Size = new System.Drawing.Size(200, 61);
            this.buttonsPanel.TabIndex = 2;
            // 
            // TodoListPage
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Name = "TodoListPage";
            this.Footer.ResumeLayout(false);
            this.Main.ResumeLayout(false);
            this.controlsPanel.ResumeLayout(false);
            this.controlsPanel.PerformLayout();
            this.buttonsPanel.ResumeLayout(false);
            this.buttonsPanel.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private Controls.IconButton btnCancel;
        private Controls.IconButton btnRefresh;
        private System.Windows.Forms.Label lblInstruction;
        private System.Windows.Forms.TableLayoutPanel controlsPanel;
        private System.Windows.Forms.TableLayoutPanel buttonsPanel;
        private System.Windows.Forms.Label lblNothingTodo;
        private Controls.LinkButton todoLinkButton;
    }
}
