namespace Fusion
{
	partial class LoginForm
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
			this.detailsButton = new Infragistics.Win.Misc.UltraButton();
			this.cancelButton = new Infragistics.Win.Misc.UltraButton();
			this.okButton = new Infragistics.Win.Misc.UltraButton();
			this.usernameLabel = new Infragistics.Win.Misc.UltraLabel();
			this.passwordLabel = new Infragistics.Win.Misc.UltraLabel();
			this.serverLabel = new Infragistics.Win.Misc.UltraLabel();
			this.databaseLabel = new Infragistics.Win.Misc.UltraLabel();
			this.usernameEditor = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.databaseEditor = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.passwordEditor = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.serverEditor = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.useIntegratedEditor = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.versionLabel = new Infragistics.Win.Misc.UltraLabel();
			this.pictureBox1 = new System.Windows.Forms.PictureBox();
			((System.ComponentModel.ISupportInitialize)(this.usernameEditor)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.databaseEditor)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.passwordEditor)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.serverEditor)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.useIntegratedEditor)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
			this.SuspendLayout();
			// 
			// detailsButton
			// 
			this.detailsButton.Location = new System.Drawing.Point(304, 194);
			this.detailsButton.Name = "detailsButton";
			this.detailsButton.Size = new System.Drawing.Size(83, 26);
			this.detailsButton.TabIndex = 11;
			this.detailsButton.Text = "&Details <<";
			// 
			// cancelButton
			// 
			this.cancelButton.DialogResult = System.Windows.Forms.DialogResult.Cancel;
			this.cancelButton.Location = new System.Drawing.Point(304, 146);
			this.cancelButton.Name = "cancelButton";
			this.cancelButton.Size = new System.Drawing.Size(83, 26);
			this.cancelButton.TabIndex = 10;
			this.cancelButton.Text = "&Cancel";
			// 
			// okButton
			// 
			this.okButton.DialogResult = System.Windows.Forms.DialogResult.OK;
			this.okButton.Location = new System.Drawing.Point(304, 114);
			this.okButton.Name = "okButton";
			this.okButton.Size = new System.Drawing.Size(83, 26);
			this.okButton.TabIndex = 9;
			this.okButton.Text = "&OK";
			// 
			// usernameLabel
			// 
			this.usernameLabel.AutoSize = true;
			this.usernameLabel.Location = new System.Drawing.Point(12, 118);
			this.usernameLabel.Name = "usernameLabel";
			this.usernameLabel.Size = new System.Drawing.Size(60, 14);
			this.usernameLabel.TabIndex = 0;
			this.usernameLabel.Text = "Username:";
			// 
			// passwordLabel
			// 
			this.passwordLabel.AutoSize = true;
			this.passwordLabel.Location = new System.Drawing.Point(12, 145);
			this.passwordLabel.Name = "passwordLabel";
			this.passwordLabel.Size = new System.Drawing.Size(57, 14);
			this.passwordLabel.TabIndex = 2;
			this.passwordLabel.Text = "Password:";
			// 
			// serverLabel
			// 
			this.serverLabel.AutoSize = true;
			this.serverLabel.Location = new System.Drawing.Point(12, 225);
			this.serverLabel.Name = "serverLabel";
			this.serverLabel.Size = new System.Drawing.Size(41, 14);
			this.serverLabel.TabIndex = 7;
			this.serverLabel.Text = "Server:";
			// 
			// databaseLabel
			// 
			this.databaseLabel.AutoSize = true;
			this.databaseLabel.Location = new System.Drawing.Point(12, 198);
			this.databaseLabel.Name = "databaseLabel";
			this.databaseLabel.Size = new System.Drawing.Size(56, 14);
			this.databaseLabel.TabIndex = 5;
			this.databaseLabel.Text = "Database:";
			// 
			// usernameEditor
			// 
			this.usernameEditor.Location = new System.Drawing.Point(80, 114);
			this.usernameEditor.Name = "usernameEditor";
			this.usernameEditor.Size = new System.Drawing.Size(210, 21);
			this.usernameEditor.TabIndex = 1;
			// 
			// databaseEditor
			// 
			this.databaseEditor.Location = new System.Drawing.Point(80, 194);
			this.databaseEditor.Name = "databaseEditor";
			this.databaseEditor.Size = new System.Drawing.Size(210, 21);
			this.databaseEditor.TabIndex = 6;
			// 
			// passwordEditor
			// 
			this.passwordEditor.Location = new System.Drawing.Point(80, 141);
			this.passwordEditor.Name = "passwordEditor";
			this.passwordEditor.PasswordChar = '*';
			this.passwordEditor.Size = new System.Drawing.Size(210, 21);
			this.passwordEditor.TabIndex = 3;
			// 
			// serverEditor
			// 
			this.serverEditor.Location = new System.Drawing.Point(81, 221);
			this.serverEditor.Name = "serverEditor";
			this.serverEditor.Size = new System.Drawing.Size(209, 21);
			this.serverEditor.TabIndex = 8;
			// 
			// useIntegratedEditor
			// 
			this.useIntegratedEditor.Location = new System.Drawing.Point(12, 168);
			this.useIntegratedEditor.Name = "useIntegratedEditor";
			this.useIntegratedEditor.Size = new System.Drawing.Size(196, 20);
			this.useIntegratedEditor.TabIndex = 4;
			this.useIntegratedEditor.Text = "&Use Windows Authentication";
			this.useIntegratedEditor.UseMnemonics = true;
			this.useIntegratedEditor.CheckedChanged += new System.EventHandler(this.UseIntegratedEditorCheckedChanged);
			// 
			// versionLabel
			// 
			this.versionLabel.AutoSize = true;
			this.versionLabel.Location = new System.Drawing.Point(8, 84);
			this.versionLabel.Name = "versionLabel";
			this.versionLabel.Size = new System.Drawing.Size(55, 14);
			this.versionLabel.TabIndex = 12;
			this.versionLabel.Text = "VERSION";
			// 
			// pictureBox1
			// 
			this.pictureBox1.Image = global::Fusion.Properties.Resources.Splash;
			this.pictureBox1.Location = new System.Drawing.Point(9, 9);
			this.pictureBox1.Name = "pictureBox1";
			this.pictureBox1.Size = new System.Drawing.Size(209, 70);
			this.pictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize;
			this.pictureBox1.TabIndex = 13;
			this.pictureBox1.TabStop = false;
			// 
			// LoginForm
			// 
			this.AcceptButton = this.okButton;
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.CancelButton = this.cancelButton;
			this.ClientSize = new System.Drawing.Size(398, 249);
			this.Controls.Add(this.databaseEditor);
			this.Controls.Add(this.serverLabel);
			this.Controls.Add(this.databaseLabel);
			this.Controls.Add(this.pictureBox1);
			this.Controls.Add(this.serverEditor);
			this.Controls.Add(this.versionLabel);
			this.Controls.Add(this.useIntegratedEditor);
			this.Controls.Add(this.passwordEditor);
			this.Controls.Add(this.usernameEditor);
			this.Controls.Add(this.passwordLabel);
			this.Controls.Add(this.usernameLabel);
			this.Controls.Add(this.okButton);
			this.Controls.Add(this.cancelButton);
			this.Controls.Add(this.detailsButton);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
			this.MaximizeBox = false;
			this.MinimizeBox = false;
			this.Name = "LoginForm";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "TITLE - Login";
			this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.LoginFormFormClosing);
			this.Load += new System.EventHandler(this.LoginFormLoad);
			((System.ComponentModel.ISupportInitialize)(this.usernameEditor)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.databaseEditor)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.passwordEditor)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.serverEditor)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.useIntegratedEditor)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
			this.ResumeLayout(false);
			this.PerformLayout();

		}

		#endregion

		private Infragistics.Win.Misc.UltraButton detailsButton;
		private Infragistics.Win.Misc.UltraButton cancelButton;
		private Infragistics.Win.Misc.UltraButton okButton;
		private Infragistics.Win.Misc.UltraLabel usernameLabel;
		private Infragistics.Win.Misc.UltraLabel passwordLabel;
		private Infragistics.Win.Misc.UltraLabel serverLabel;
		private Infragistics.Win.Misc.UltraLabel databaseLabel;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor usernameEditor;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor databaseEditor;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor passwordEditor;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor serverEditor;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor useIntegratedEditor;
		private Infragistics.Win.Misc.UltraLabel versionLabel;
		private System.Windows.Forms.PictureBox pictureBox1;
	}
}