namespace Fusion.Pages
{
	partial class Settings
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
			this.ultraLabel1 = new Infragistics.Win.Misc.UltraLabel();
			this.ultraLabel2 = new Infragistics.Win.Misc.UltraLabel();
			this.ultraLabel3 = new Infragistics.Win.Misc.UltraLabel();
			this.ultraLabel4 = new Infragistics.Win.Misc.UltraLabel();
			this.communityDatabase = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.serviceStatusLabel = new System.Windows.Forms.Label();
			((System.ComponentModel.ISupportInitialize)(this.communityDatabase)).BeginInit();
			this.SuspendLayout();
			// 
			// ultraLabel1
			// 
			this.ultraLabel1.Anchor = System.Windows.Forms.AnchorStyles.Top;
			this.ultraLabel1.AutoSize = true;
			this.ultraLabel1.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.ultraLabel1.Location = new System.Drawing.Point(123, 24);
			this.ultraLabel1.Name = "ultraLabel1";
			this.ultraLabel1.Size = new System.Drawing.Size(387, 17);
			this.ultraLabel1.TabIndex = 0;
			this.ultraLabel1.Text = "Please enter the required community string for this database";
			// 
			// ultraLabel2
			// 
			this.ultraLabel2.Anchor = System.Windows.Forms.AnchorStyles.Top;
			this.ultraLabel2.AutoSize = true;
			this.ultraLabel2.Location = new System.Drawing.Point(141, 57);
			this.ultraLabel2.Name = "ultraLabel2";
			this.ultraLabel2.Size = new System.Drawing.Size(361, 14);
			this.ultraLabel2.TabIndex = 1;
			this.ultraLabel2.Text = "This should be in the format <customer>.<environment>.<database>db";
			// 
			// ultraLabel3
			// 
			this.ultraLabel3.Anchor = System.Windows.Forms.AnchorStyles.Top;
			this.ultraLabel3.AutoSize = true;
			this.ultraLabel3.Location = new System.Drawing.Point(133, 134);
			this.ultraLabel3.Name = "ultraLabel3";
			this.ultraLabel3.Size = new System.Drawing.Size(376, 14);
			this.ultraLabel3.TabIndex = 2;
			this.ultraLabel3.Text = "e.g. advanced.dev.livedb, advanced.live.livedb or advanced.live.trainingdb";
			// 
			// ultraLabel4
			// 
			this.ultraLabel4.Anchor = System.Windows.Forms.AnchorStyles.Top;
			this.ultraLabel4.AutoSize = true;
			this.ultraLabel4.Location = new System.Drawing.Point(49, 169);
			this.ultraLabel4.Name = "ultraLabel4";
			this.ultraLabel4.Size = new System.Drawing.Size(535, 14);
			this.ultraLabel4.TabIndex = 3;
			this.ultraLabel4.Text = "This will be passed as part of each message and must be the same for all systems " +
    "connected using fusion";
			// 
			// communityDatabase
			// 
			this.communityDatabase.Anchor = System.Windows.Forms.AnchorStyles.Top;
			this.communityDatabase.Location = new System.Drawing.Point(133, 92);
			this.communityDatabase.Name = "communityDatabase";
			this.communityDatabase.Size = new System.Drawing.Size(369, 21);
			this.communityDatabase.TabIndex = 4;
			// 
			// serviceStatusLabel
			// 
			this.serviceStatusLabel.Anchor = System.Windows.Forms.AnchorStyles.Top;
			this.serviceStatusLabel.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(255)))), ((int)(((byte)(192)))));
			this.serviceStatusLabel.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
			this.serviceStatusLabel.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.serviceStatusLabel.Location = new System.Drawing.Point(106, 204);
			this.serviceStatusLabel.Name = "serviceStatusLabel";
			this.serviceStatusLabel.Size = new System.Drawing.Size(420, 33);
			this.serviceStatusLabel.TabIndex = 6;
			this.serviceStatusLabel.Text = "The Fusion Connector Service is not installed on this computer";
			this.serviceStatusLabel.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// Settings
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.Controls.Add(this.serviceStatusLabel);
			this.Controls.Add(this.communityDatabase);
			this.Controls.Add(this.ultraLabel4);
			this.Controls.Add(this.ultraLabel3);
			this.Controls.Add(this.ultraLabel2);
			this.Controls.Add(this.ultraLabel1);
			this.Name = "Settings";
			this.Size = new System.Drawing.Size(632, 310);
			((System.ComponentModel.ISupportInitialize)(this.communityDatabase)).EndInit();
			this.ResumeLayout(false);
			this.PerformLayout();

		}

		#endregion

		private Infragistics.Win.Misc.UltraLabel ultraLabel1;
		private Infragistics.Win.Misc.UltraLabel ultraLabel2;
		private Infragistics.Win.Misc.UltraLabel ultraLabel3;
		private Infragistics.Win.Misc.UltraLabel ultraLabel4;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor communityDatabase;
		private System.Windows.Forms.Label serviceStatusLabel;
	}
}
