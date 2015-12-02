namespace MobileDesigner.Pages
{
    partial class NewRegistrationPage
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
            this.txtEmail = new System.Windows.Forms.TextBox();
            this.lblEmail = new System.Windows.Forms.Label();
            this.btnCancel = new MobileDesigner.Controls.IconButton();
            this.btnSubmit = new MobileDesigner.Controls.IconButton();
            this.lblWelcome = new System.Windows.Forms.Label();
            this.controlsPanel = new System.Windows.Forms.TableLayoutPanel();
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
            this.Footer.TabIndex = 10;
            // 
            // Main
            // 
            this.Main.Controls.Add(this.controlsPanel);
            // 
            // txtEmail
            // 
            this.txtEmail.Dock = System.Windows.Forms.DockStyle.Top;
            this.txtEmail.Location = new System.Drawing.Point(84, 16);
            this.txtEmail.Name = "txtEmail";
            this.txtEmail.Size = new System.Drawing.Size(161, 20);
            this.txtEmail.TabIndex = 3;
            // 
            // lblEmail
            // 
            this.lblEmail.Anchor = System.Windows.Forms.AnchorStyles.Left;
            this.lblEmail.AutoSize = true;
            this.lblEmail.Location = new System.Drawing.Point(3, 19);
            this.lblEmail.Name = "lblEmail";
            this.lblEmail.Size = new System.Drawing.Size(75, 13);
            this.lblEmail.TabIndex = 2;
            this.lblEmail.Text = "Email address:";
            this.lblEmail.UseMnemonic = false;
            // 
            // btnCancel
            // 
            this.btnCancel.Anchor = System.Windows.Forms.AnchorStyles.Top;
            this.btnCancel.AutoSize = true;
            this.btnCancel.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.btnCancel.Caption = "Cancel";
            this.btnCancel.Image = null;
            this.btnCancel.Location = new System.Drawing.Point(130, 3);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(32, 50);
            this.btnCancel.TabIndex = 1;
            // 
            // btnSubmit
            // 
            this.btnSubmit.Anchor = System.Windows.Forms.AnchorStyles.Top;
            this.btnSubmit.AutoSize = true;
            this.btnSubmit.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.btnSubmit.Caption = "OK";
            this.btnSubmit.Image = null;
            this.btnSubmit.Location = new System.Drawing.Point(32, 3);
            this.btnSubmit.Name = "btnSubmit";
            this.btnSubmit.Size = new System.Drawing.Size(32, 50);
            this.btnSubmit.TabIndex = 0;
            // 
            // lblWelcome
            // 
            this.lblWelcome.AutoSize = true;
            this.controlsPanel.SetColumnSpan(this.lblWelcome, 2);
            this.lblWelcome.Location = new System.Drawing.Point(0, 0);
            this.lblWelcome.Margin = new System.Windows.Forms.Padding(0);
            this.lblWelcome.Name = "lblWelcome";
            this.lblWelcome.Size = new System.Drawing.Size(91, 13);
            this.lblWelcome.TabIndex = 0;
            this.lblWelcome.Text = "Welcome Caption";
            this.lblWelcome.UseMnemonic = false;
            // 
            // controlsPanel
            // 
            this.controlsPanel.ColumnCount = 2;
            this.controlsPanel.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle());
            this.controlsPanel.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.controlsPanel.Controls.Add(this.lblEmail, 0, 1);
            this.controlsPanel.Controls.Add(this.txtEmail, 1, 1);
            this.controlsPanel.Controls.Add(this.lblWelcome, 0, 0);
            this.controlsPanel.Location = new System.Drawing.Point(0, 28);
            this.controlsPanel.Name = "controlsPanel";
            this.controlsPanel.RowCount = 3;
            this.controlsPanel.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.controlsPanel.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.controlsPanel.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20F));
            this.controlsPanel.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20F));
            this.controlsPanel.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20F));
            this.controlsPanel.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20F));
            this.controlsPanel.Size = new System.Drawing.Size(248, 156);
            this.controlsPanel.TabIndex = 10;
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
            this.buttonsPanel.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.buttonsPanel.Size = new System.Drawing.Size(195, 51);
            this.buttonsPanel.TabIndex = 11;
            // 
            // NewRegistrationPage
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Name = "NewRegistrationPage";
            this.Footer.ResumeLayout(false);
            this.Main.ResumeLayout(false);
            this.controlsPanel.ResumeLayout(false);
            this.controlsPanel.PerformLayout();
            this.buttonsPanel.ResumeLayout(false);
            this.buttonsPanel.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TextBox txtEmail;
        private System.Windows.Forms.Label lblEmail;
        private Controls.IconButton btnSubmit;
        private Controls.IconButton btnCancel;
        private System.Windows.Forms.Label lblWelcome;
        private System.Windows.Forms.TableLayoutPanel controlsPanel;
        private System.Windows.Forms.TableLayoutPanel buttonsPanel;
    }
}
