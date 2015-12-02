namespace MobileDesigner.Pages
{
    partial class LoginPage
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
            this.btnForgotPwd = new MobileDesigner.Controls.IconButton();
            this.btnRegister = new MobileDesigner.Controls.IconButton();
            this.btnLogin = new MobileDesigner.Controls.IconButton();
            this.label1 = new System.Windows.Forms.Label();
            this.controlsPanel = new System.Windows.Forms.TableLayoutPanel();
            this.chkRememberPwd = new System.Windows.Forms.CheckBox();
            this.lblRememberPwd = new System.Windows.Forms.Label();
            this.lblPassword = new System.Windows.Forms.Label();
            this.lblUserName = new System.Windows.Forms.Label();
            this.txtPassword = new System.Windows.Forms.TextBox();
            this.txtUserName = new System.Windows.Forms.TextBox();
            this.lblWelcome = new System.Windows.Forms.Label();
            this.buttonsPanel = new System.Windows.Forms.TableLayoutPanel();
            this.iconButton2 = new MobileDesigner.Controls.IconButton();
            this.copyRightLabel = new System.Windows.Forms.Label();
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
            this.Main.BackgroundImageLayout = MobileDesigner.ImageLayout.TopRight;
            this.Main.Controls.Add(this.copyRightLabel);
            this.Main.Controls.Add(this.controlsPanel);
            // 
            // btnForgotPwd
            // 
            this.btnForgotPwd.Anchor = System.Windows.Forms.AnchorStyles.Top;
            this.btnForgotPwd.AutoSize = true;
            this.btnForgotPwd.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.btnForgotPwd.Caption = "Forgot login?";
            this.btnForgotPwd.Image = null;
            this.btnForgotPwd.Location = new System.Drawing.Point(96, 3);
            this.btnForgotPwd.Name = "btnForgotPwd";
            this.btnForgotPwd.Size = new System.Drawing.Size(57, 50);
            this.btnForgotPwd.TabIndex = 1;
            // 
            // btnRegister
            // 
            this.btnRegister.Anchor = System.Windows.Forms.AnchorStyles.Top;
            this.btnRegister.AutoSize = true;
            this.btnRegister.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.btnRegister.Caption = "Register";
            this.btnRegister.Image = null;
            this.btnRegister.Location = new System.Drawing.Point(189, 3);
            this.btnRegister.Name = "btnRegister";
            this.btnRegister.Size = new System.Drawing.Size(38, 50);
            this.btnRegister.TabIndex = 2;
            // 
            // btnLogin
            // 
            this.btnLogin.Anchor = System.Windows.Forms.AnchorStyles.Top;
            this.btnLogin.AutoSize = true;
            this.btnLogin.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.btnLogin.Caption = "Login";
            this.btnLogin.Image = null;
            this.btnLogin.Location = new System.Drawing.Point(25, 3);
            this.btnLogin.Name = "btnLogin";
            this.btnLogin.Size = new System.Drawing.Size(32, 50);
            this.btnLogin.TabIndex = 0;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(32, 96);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(35, 13);
            this.label1.TabIndex = 1;
            this.label1.Text = "label1";
            // 
            // controlsPanel
            // 
            this.controlsPanel.ColumnCount = 2;
            this.controlsPanel.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle());
            this.controlsPanel.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.controlsPanel.Controls.Add(this.chkRememberPwd, 1, 3);
            this.controlsPanel.Controls.Add(this.lblRememberPwd, 0, 3);
            this.controlsPanel.Controls.Add(this.lblPassword, 0, 2);
            this.controlsPanel.Controls.Add(this.lblUserName, 0, 1);
            this.controlsPanel.Controls.Add(this.txtPassword, 1, 2);
            this.controlsPanel.Controls.Add(this.txtUserName, 1, 1);
            this.controlsPanel.Controls.Add(this.lblWelcome, 0, 0);
            this.controlsPanel.Location = new System.Drawing.Point(21, 19);
            this.controlsPanel.Name = "controlsPanel";
            this.controlsPanel.RowCount = 5;
            this.controlsPanel.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.controlsPanel.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.controlsPanel.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.controlsPanel.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.controlsPanel.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20F));
            this.controlsPanel.Size = new System.Drawing.Size(244, 161);
            this.controlsPanel.TabIndex = 3;
            // 
            // chkRememberPwd
            // 
            this.chkRememberPwd.AutoSize = true;
            this.chkRememberPwd.Location = new System.Drawing.Point(106, 68);
            this.chkRememberPwd.Name = "chkRememberPwd";
            this.chkRememberPwd.Size = new System.Drawing.Size(15, 14);
            this.chkRememberPwd.TabIndex = 6;
            this.chkRememberPwd.UseVisualStyleBackColor = true;
            // 
            // lblRememberPwd
            // 
            this.lblRememberPwd.Anchor = System.Windows.Forms.AnchorStyles.Left;
            this.lblRememberPwd.AutoSize = true;
            this.lblRememberPwd.Location = new System.Drawing.Point(3, 68);
            this.lblRememberPwd.Name = "lblRememberPwd";
            this.lblRememberPwd.Size = new System.Drawing.Size(97, 13);
            this.lblRememberPwd.TabIndex = 5;
            this.lblRememberPwd.Text = "Keep me signed in:";
            this.lblRememberPwd.UseMnemonic = false;
            // 
            // lblPassword
            // 
            this.lblPassword.Anchor = System.Windows.Forms.AnchorStyles.Left;
            this.lblPassword.AutoSize = true;
            this.lblPassword.Location = new System.Drawing.Point(3, 45);
            this.lblPassword.Name = "lblPassword";
            this.lblPassword.Size = new System.Drawing.Size(56, 13);
            this.lblPassword.TabIndex = 3;
            this.lblPassword.Text = "Password:";
            this.lblPassword.UseMnemonic = false;
            // 
            // lblUserName
            // 
            this.lblUserName.Anchor = System.Windows.Forms.AnchorStyles.Left;
            this.lblUserName.AutoSize = true;
            this.lblUserName.Location = new System.Drawing.Point(3, 19);
            this.lblUserName.Name = "lblUserName";
            this.lblUserName.Size = new System.Drawing.Size(58, 13);
            this.lblUserName.TabIndex = 1;
            this.lblUserName.Text = "Username:";
            this.lblUserName.UseMnemonic = false;
            // 
            // txtPassword
            // 
            this.txtPassword.Dock = System.Windows.Forms.DockStyle.Top;
            this.txtPassword.Location = new System.Drawing.Point(106, 42);
            this.txtPassword.Name = "txtPassword";
            this.txtPassword.Size = new System.Drawing.Size(135, 20);
            this.txtPassword.TabIndex = 4;
            // 
            // txtUserName
            // 
            this.txtUserName.Dock = System.Windows.Forms.DockStyle.Top;
            this.txtUserName.Location = new System.Drawing.Point(106, 16);
            this.txtUserName.Name = "txtUserName";
            this.txtUserName.Size = new System.Drawing.Size(135, 20);
            this.txtUserName.TabIndex = 2;
            // 
            // lblWelcome
            // 
            this.lblWelcome.AutoSize = true;
            this.lblWelcome.BackColor = System.Drawing.Color.Transparent;
            this.controlsPanel.SetColumnSpan(this.lblWelcome, 2);
            this.lblWelcome.Location = new System.Drawing.Point(3, 0);
            this.lblWelcome.Name = "lblWelcome";
            this.lblWelcome.Size = new System.Drawing.Size(91, 13);
            this.lblWelcome.TabIndex = 0;
            this.lblWelcome.Text = "Welcome Caption";
            this.lblWelcome.UseMnemonic = false;
            // 
            // buttonsPanel
            // 
            this.buttonsPanel.ColumnCount = 3;
            this.buttonsPanel.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 33F));
            this.buttonsPanel.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 34F));
            this.buttonsPanel.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 33F));
            this.buttonsPanel.Controls.Add(this.btnLogin, 0, 0);
            this.buttonsPanel.Controls.Add(this.btnForgotPwd, 1, 0);
            this.buttonsPanel.Controls.Add(this.btnRegister, 2, 0);
            this.buttonsPanel.Location = new System.Drawing.Point(4, 4);
            this.buttonsPanel.Name = "buttonsPanel";
            this.buttonsPanel.RowCount = 1;
            this.buttonsPanel.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.buttonsPanel.Size = new System.Drawing.Size(250, 53);
            this.buttonsPanel.TabIndex = 4;
            // 
            // iconButton2
            // 
            this.iconButton2.AutoSize = true;
            this.iconButton2.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.iconButton2.Caption = "Forgot login?";
            this.iconButton2.Image = null;
            this.iconButton2.Location = new System.Drawing.Point(69, 3);
            this.iconButton2.Name = "iconButton2";
            this.iconButton2.Size = new System.Drawing.Size(40, 50);
            this.iconButton2.TabIndex = 2;
            // 
            // copyRightLabel
            // 
            this.copyRightLabel.BackColor = System.Drawing.Color.Transparent;
            this.copyRightLabel.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.copyRightLabel.Font = new System.Drawing.Font("Verdana", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Pixel, ((byte)(0)));
            this.copyRightLabel.Location = new System.Drawing.Point(0, 233);
            this.copyRightLabel.Name = "copyRightLabel";
            this.copyRightLabel.Size = new System.Drawing.Size(270, 35);
            this.copyRightLabel.TabIndex = 4;
            this.copyRightLabel.Tag = "NOSELECT";
            this.copyRightLabel.Text = "Copyright © Advanced Business Software and Solutions Ltd 2013";
            this.copyRightLabel.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            this.copyRightLabel.UseMnemonic = false;
            // 
            // LoginPage
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Name = "LoginPage";
            this.Footer.ResumeLayout(false);
            this.Main.ResumeLayout(false);
            this.controlsPanel.ResumeLayout(false);
            this.controlsPanel.PerformLayout();
            this.buttonsPanel.ResumeLayout(false);
            this.buttonsPanel.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private Controls.IconButton btnLogin;
        private Controls.IconButton btnRegister;
        private Controls.IconButton btnForgotPwd;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TableLayoutPanel controlsPanel;
        private System.Windows.Forms.Label lblRememberPwd;
        private System.Windows.Forms.Label lblUserName;
        private System.Windows.Forms.TextBox txtUserName;
        private System.Windows.Forms.TextBox txtPassword;
        private System.Windows.Forms.CheckBox chkRememberPwd;
        private System.Windows.Forms.Label lblPassword;
        private System.Windows.Forms.TableLayoutPanel buttonsPanel;
        private Controls.IconButton iconButton2;
        private System.Windows.Forms.Label lblWelcome;
        private System.Windows.Forms.Label copyRightLabel;

    }
}
