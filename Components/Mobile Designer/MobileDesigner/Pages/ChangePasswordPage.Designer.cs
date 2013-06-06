namespace MobileDesigner.Pages
{
    partial class ChangePasswordPage
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
            this.btnCancel = new MobileDesigner.Controls.IconButton();
            this.btnSubmit = new MobileDesigner.Controls.IconButton();
            this.controlsPanel = new System.Windows.Forms.TableLayoutPanel();
            this.txtConfPassword = new System.Windows.Forms.TextBox();
            this.txtCurrPassword = new System.Windows.Forms.TextBox();
            this.txtNewPassword = new System.Windows.Forms.TextBox();
            this.lblNewPassword = new System.Windows.Forms.Label();
            this.lblConfPassword = new System.Windows.Forms.Label();
            this.lblCurrPassword = new System.Windows.Forms.Label();
            this.lblWelcome = new System.Windows.Forms.Label();
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
            // btnSubmit
            // 
            this.btnSubmit.Anchor = System.Windows.Forms.AnchorStyles.Top;
            this.btnSubmit.AutoSize = true;
            this.btnSubmit.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.btnSubmit.Caption = "Submit";
            this.btnSubmit.Image = null;
            this.btnSubmit.Location = new System.Drawing.Point(33, 3);
            this.btnSubmit.Name = "btnSubmit";
            this.btnSubmit.Size = new System.Drawing.Size(33, 50);
            this.btnSubmit.TabIndex = 0;
            // 
            // controlsPanel
            // 
            this.controlsPanel.ColumnCount = 2;
            this.controlsPanel.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle());
            this.controlsPanel.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 118F));
            this.controlsPanel.Controls.Add(this.txtConfPassword, 1, 3);
            this.controlsPanel.Controls.Add(this.txtCurrPassword, 1, 1);
            this.controlsPanel.Controls.Add(this.txtNewPassword, 1, 2);
            this.controlsPanel.Controls.Add(this.lblNewPassword, 0, 2);
            this.controlsPanel.Controls.Add(this.lblConfPassword, 0, 3);
            this.controlsPanel.Controls.Add(this.lblCurrPassword, 0, 1);
            this.controlsPanel.Controls.Add(this.lblWelcome, 0, 0);
            this.controlsPanel.Location = new System.Drawing.Point(3, 26);
            this.controlsPanel.Name = "controlsPanel";
            this.controlsPanel.RowCount = 5;
            this.controlsPanel.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.controlsPanel.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.controlsPanel.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.controlsPanel.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.controlsPanel.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20F));
            this.controlsPanel.Size = new System.Drawing.Size(246, 117);
            this.controlsPanel.TabIndex = 14;
            // 
            // txtConfPassword
            // 
            this.txtConfPassword.Dock = System.Windows.Forms.DockStyle.Top;
            this.txtConfPassword.Location = new System.Drawing.Point(131, 68);
            this.txtConfPassword.Name = "txtConfPassword";
            this.txtConfPassword.Size = new System.Drawing.Size(112, 20);
            this.txtConfPassword.TabIndex = 11;
            // 
            // txtCurrPassword
            // 
            this.txtCurrPassword.Anchor = System.Windows.Forms.AnchorStyles.Left;
            this.txtCurrPassword.Location = new System.Drawing.Point(131, 16);
            this.txtCurrPassword.Name = "txtCurrPassword";
            this.txtCurrPassword.Size = new System.Drawing.Size(112, 20);
            this.txtCurrPassword.TabIndex = 7;
            // 
            // txtNewPassword
            // 
            this.txtNewPassword.Dock = System.Windows.Forms.DockStyle.Top;
            this.txtNewPassword.Location = new System.Drawing.Point(131, 42);
            this.txtNewPassword.Name = "txtNewPassword";
            this.txtNewPassword.Size = new System.Drawing.Size(112, 20);
            this.txtNewPassword.TabIndex = 9;
            // 
            // lblNewPassword
            // 
            this.lblNewPassword.Anchor = System.Windows.Forms.AnchorStyles.Left;
            this.lblNewPassword.AutoSize = true;
            this.lblNewPassword.Location = new System.Drawing.Point(3, 45);
            this.lblNewPassword.Name = "lblNewPassword";
            this.lblNewPassword.Size = new System.Drawing.Size(81, 13);
            this.lblNewPassword.TabIndex = 8;
            this.lblNewPassword.Text = "New Password:";
            this.lblNewPassword.UseMnemonic = false;
            // 
            // lblConfPassword
            // 
            this.lblConfPassword.Anchor = System.Windows.Forms.AnchorStyles.Left;
            this.lblConfPassword.AutoSize = true;
            this.lblConfPassword.Location = new System.Drawing.Point(3, 71);
            this.lblConfPassword.Name = "lblConfPassword";
            this.lblConfPassword.Size = new System.Drawing.Size(122, 13);
            this.lblConfPassword.TabIndex = 10;
            this.lblConfPassword.Text = "Re-enter new password:";
            this.lblConfPassword.UseMnemonic = false;
            // 
            // lblCurrPassword
            // 
            this.lblCurrPassword.Anchor = System.Windows.Forms.AnchorStyles.Left;
            this.lblCurrPassword.AutoSize = true;
            this.lblCurrPassword.Location = new System.Drawing.Point(3, 19);
            this.lblCurrPassword.Name = "lblCurrPassword";
            this.lblCurrPassword.Size = new System.Drawing.Size(92, 13);
            this.lblCurrPassword.TabIndex = 6;
            this.lblCurrPassword.Text = "Current password:";
            this.lblCurrPassword.UseMnemonic = false;
            // 
            // lblWelcome
            // 
            this.lblWelcome.AutoSize = true;
            this.controlsPanel.SetColumnSpan(this.lblWelcome, 2);
            this.lblWelcome.Location = new System.Drawing.Point(3, 0);
            this.lblWelcome.Name = "lblWelcome";
            this.lblWelcome.Size = new System.Drawing.Size(91, 13);
            this.lblWelcome.TabIndex = 12;
            this.lblWelcome.Text = "Welcome Caption";
            this.lblWelcome.UseMnemonic = false;
            // 
            // buttonsPanel
            // 
            this.buttonsPanel.ColumnCount = 2;
            this.buttonsPanel.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.buttonsPanel.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.buttonsPanel.Controls.Add(this.btnSubmit, 0, 0);
            this.buttonsPanel.Controls.Add(this.btnCancel, 1, 0);
            this.buttonsPanel.Location = new System.Drawing.Point(4, 4);
            this.buttonsPanel.Name = "buttonsPanel";
            this.buttonsPanel.RowCount = 1;
            this.buttonsPanel.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.buttonsPanel.Size = new System.Drawing.Size(200, 61);
            this.buttonsPanel.TabIndex = 15;
            // 
            // ChangePasswordPage
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Name = "ChangePasswordPage";
            this.Footer.ResumeLayout(false);
            this.Main.ResumeLayout(false);
            this.controlsPanel.ResumeLayout(false);
            this.controlsPanel.PerformLayout();
            this.buttonsPanel.ResumeLayout(false);
            this.buttonsPanel.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private Controls.IconButton btnSubmit;
        private Controls.IconButton btnCancel;
        private System.Windows.Forms.TableLayoutPanel controlsPanel;
        private System.Windows.Forms.TextBox txtConfPassword;
        private System.Windows.Forms.TextBox txtCurrPassword;
        private System.Windows.Forms.TextBox txtNewPassword;
        private System.Windows.Forms.Label lblNewPassword;
        private System.Windows.Forms.Label lblConfPassword;
        private System.Windows.Forms.Label lblCurrPassword;
        private System.Windows.Forms.TableLayoutPanel buttonsPanel;
        private System.Windows.Forms.Label lblWelcome;
    }
}
