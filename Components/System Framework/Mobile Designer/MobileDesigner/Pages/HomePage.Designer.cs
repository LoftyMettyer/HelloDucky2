namespace MobileDesigner.Pages
{
    partial class HomePage
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
            this.btnLogout = new MobileDesigner.Controls.IconButton();
            this.btnChangePwd = new MobileDesigner.Controls.IconButton();
            this.buttonsPanel = new System.Windows.Forms.TableLayoutPanel();
            this.btnToDoList = new MobileDesigner.Controls.IconButton();
            this.controlsPanel = new System.Windows.Forms.TableLayoutPanel();
            this.ItemsList = new MobileDesigner.Controls.ItemsList();
            this.lblWelcome = new System.Windows.Forms.Label();
            this.lblNothingTodo = new System.Windows.Forms.Label();
            this.Footer.SuspendLayout();
            this.Main.SuspendLayout();
            this.buttonsPanel.SuspendLayout();
            this.controlsPanel.SuspendLayout();
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
            // btnLogout
            // 
            this.btnLogout.Anchor = System.Windows.Forms.AnchorStyles.Top;
            this.btnLogout.AutoSize = true;
            this.btnLogout.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.btnLogout.Caption = "Logout";
            this.btnLogout.Image = null;
            this.btnLogout.Location = new System.Drawing.Point(203, 3);
            this.btnLogout.Name = "btnLogout";
            this.btnLogout.Size = new System.Drawing.Size(32, 50);
            this.btnLogout.TabIndex = 2;
            // 
            // btnChangePwd
            // 
            this.btnChangePwd.Anchor = System.Windows.Forms.AnchorStyles.Top;
            this.btnChangePwd.AutoSize = true;
            this.btnChangePwd.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.btnChangePwd.Caption = "Change Password";
            this.btnChangePwd.Image = null;
            this.btnChangePwd.Location = new System.Drawing.Point(92, 3);
            this.btnChangePwd.Name = "btnChangePwd";
            this.btnChangePwd.Size = new System.Drawing.Size(77, 50);
            this.btnChangePwd.TabIndex = 1;
            // 
            // buttonsPanel
            // 
            this.buttonsPanel.ColumnCount = 3;
            this.buttonsPanel.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 33F));
            this.buttonsPanel.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 34F));
            this.buttonsPanel.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 33F));
            this.buttonsPanel.Controls.Add(this.btnToDoList, 0, 0);
            this.buttonsPanel.Controls.Add(this.btnLogout, 2, 0);
            this.buttonsPanel.Controls.Add(this.btnChangePwd, 1, 0);
            this.buttonsPanel.Location = new System.Drawing.Point(4, 4);
            this.buttonsPanel.Name = "buttonsPanel";
            this.buttonsPanel.RowCount = 1;
            this.buttonsPanel.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.buttonsPanel.Size = new System.Drawing.Size(263, 68);
            this.buttonsPanel.TabIndex = 0;
            // 
            // btnToDoList
            // 
            this.btnToDoList.Anchor = System.Windows.Forms.AnchorStyles.Top;
            this.btnToDoList.AutoSize = true;
            this.btnToDoList.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.btnToDoList.Caption = "Todo List";
            this.btnToDoList.Image = null;
            this.btnToDoList.Location = new System.Drawing.Point(22, 3);
            this.btnToDoList.Name = "btnToDoList";
            this.btnToDoList.Size = new System.Drawing.Size(41, 50);
            this.btnToDoList.TabIndex = 0;
            // 
            // controlsPanel
            // 
            this.controlsPanel.ColumnCount = 2;
            this.controlsPanel.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle());
            this.controlsPanel.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.controlsPanel.Controls.Add(this.ItemsList, 0, 2);
            this.controlsPanel.Controls.Add(this.lblWelcome, 0, 1);
            this.controlsPanel.Controls.Add(this.lblNothingTodo, 0, 0);
            this.controlsPanel.Location = new System.Drawing.Point(8, 6);
            this.controlsPanel.Name = "controlsPanel";
            this.controlsPanel.RowCount = 3;
            this.controlsPanel.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.controlsPanel.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.controlsPanel.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.controlsPanel.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20F));
            this.controlsPanel.Size = new System.Drawing.Size(258, 256);
            this.controlsPanel.TabIndex = 0;
            // 
            // ItemsList
            // 
            this.controlsPanel.SetColumnSpan(this.ItemsList, 2);
            this.ItemsList.Dock = System.Windows.Forms.DockStyle.Fill;
            this.ItemsList.Font = new System.Drawing.Font("Verdana", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.ItemsList.ItemFont = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.ItemsList.ItemForeColor = System.Drawing.SystemColors.ControlText;
            this.ItemsList.Location = new System.Drawing.Point(3, 29);
            this.ItemsList.Name = "ItemsList";
            this.ItemsList.Padding = new System.Windows.Forms.Padding(3);
            this.ItemsList.SelectedUserGroup = null;
            this.ItemsList.SelectedWorkflows = null;
            this.ItemsList.Size = new System.Drawing.Size(252, 224);
            this.ItemsList.TabIndex = 1;
            this.ItemsList.UserGroups = null;
            this.ItemsList.Workflows = null;
            // 
            // lblWelcome
            // 
            this.lblWelcome.AutoSize = true;
            this.lblWelcome.BackColor = System.Drawing.Color.Transparent;
            this.controlsPanel.SetColumnSpan(this.lblWelcome, 2);
            this.lblWelcome.Location = new System.Drawing.Point(3, 13);
            this.lblWelcome.Name = "lblWelcome";
            this.lblWelcome.Size = new System.Drawing.Size(91, 13);
            this.lblWelcome.TabIndex = 0;
            this.lblWelcome.Text = "Welcome Caption";
            this.lblWelcome.UseMnemonic = false;
            // 
            // lblNothingTodo
            // 
            this.lblNothingTodo.AutoSize = true;
            this.lblNothingTodo.BackColor = System.Drawing.Color.Transparent;
            this.controlsPanel.SetColumnSpan(this.lblNothingTodo, 2);
            this.lblNothingTodo.Location = new System.Drawing.Point(3, 0);
            this.lblNothingTodo.Name = "lblNothingTodo";
            this.lblNothingTodo.Size = new System.Drawing.Size(165, 13);
            this.lblNothingTodo.TabIndex = 3;
            this.lblNothingTodo.Text = "Welcome Caption - no items in list";
            this.lblNothingTodo.UseMnemonic = false;
            // 
            // HomePage
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Name = "HomePage";
            this.Footer.ResumeLayout(false);
            this.Main.ResumeLayout(false);
            this.buttonsPanel.ResumeLayout(false);
            this.buttonsPanel.PerformLayout();
            this.controlsPanel.ResumeLayout(false);
            this.controlsPanel.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private Controls.IconButton btnChangePwd;
        private Controls.IconButton btnLogout;
        private System.Windows.Forms.TableLayoutPanel buttonsPanel;
        private Controls.IconButton btnToDoList;
        private System.Windows.Forms.TableLayoutPanel controlsPanel;
        private System.Windows.Forms.Label lblWelcome;
        public Controls.ItemsList ItemsList;
        private System.Windows.Forms.Label lblNothingTodo;
    }
}
