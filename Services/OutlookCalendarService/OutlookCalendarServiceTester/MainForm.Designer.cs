namespace OutlookCalendarService2Tester
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
	  System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MainForm));
	  this.btnCustomer = new System.Windows.Forms.Button();
	  this.btnSupport = new System.Windows.Forms.Button();
	  this.txtSupportPassword = new System.Windows.Forms.TextBox();
	  this.label1 = new System.Windows.Forms.Label();
	  this.label2 = new System.Windows.Forms.Label();
	  this.SuspendLayout();
	  // 
	  // btnCustomer
	  // 
	  this.btnCustomer.Location = new System.Drawing.Point(211, 94);
	  this.btnCustomer.Name = "btnCustomer";
	  this.btnCustomer.Size = new System.Drawing.Size(100, 23);
	  this.btnCustomer.TabIndex = 3;
	  this.btnCustomer.Text = "Customer";
	  this.btnCustomer.UseVisualStyleBackColor = true;
	  this.btnCustomer.Click += new System.EventHandler(this.btnCustomer_Click);
	  // 
	  // btnSupport
	  // 
	  this.btnSupport.Location = new System.Drawing.Point(73, 94);
	  this.btnSupport.Name = "btnSupport";
	  this.btnSupport.Size = new System.Drawing.Size(100, 23);
	  this.btnSupport.TabIndex = 1;
	  this.btnSupport.Text = "Support";
	  this.btnSupport.UseVisualStyleBackColor = true;
	  this.btnSupport.Click += new System.EventHandler(this.btnSupport_Click);
	  // 
	  // txtSupportPassword
	  // 
	  this.txtSupportPassword.Location = new System.Drawing.Point(73, 157);
	  this.txtSupportPassword.Name = "txtSupportPassword";
	  this.txtSupportPassword.PasswordChar = '*';
	  this.txtSupportPassword.Size = new System.Drawing.Size(100, 20);
	  this.txtSupportPassword.TabIndex = 2;
	  this.txtSupportPassword.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtSupportPassword_KeyPress);
	  // 
	  // label1
	  // 
	  this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
	  this.label1.Location = new System.Drawing.Point(19, 30);
	  this.label1.Name = "label1";
	  this.label1.Size = new System.Drawing.Size(347, 41);
	  this.label1.TabIndex = 3;
	  this.label1.Text = "OpenHR Outlook Calendar Tester";
	  this.label1.TextAlign = System.Drawing.ContentAlignment.TopCenter;
	  // 
	  // label2
	  // 
	  this.label2.Location = new System.Drawing.Point(70, 141);
	  this.label2.Name = "label2";
	  this.label2.Size = new System.Drawing.Size(100, 13);
	  this.label2.TabIndex = 4;
	  this.label2.Text = "Password";
	  this.label2.TextAlign = System.Drawing.ContentAlignment.TopCenter;
	  // 
	  // MainForm
	  // 
	  this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
	  this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
	  this.ClientSize = new System.Drawing.Size(385, 194);
	  this.Controls.Add(this.label2);
	  this.Controls.Add(this.label1);
	  this.Controls.Add(this.txtSupportPassword);
	  this.Controls.Add(this.btnSupport);
	  this.Controls.Add(this.btnCustomer);
	  this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
	  this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
	  this.MaximizeBox = false;
	  this.Name = "MainForm";
	  this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
	  this.Text = "OpenHR";
	  this.ResumeLayout(false);
	  this.PerformLayout();

	}

	#endregion

	private System.Windows.Forms.Button btnCustomer;
	private System.Windows.Forms.Button btnSupport;
	private System.Windows.Forms.TextBox txtSupportPassword;
	private System.Windows.Forms.Label label1;
	private System.Windows.Forms.Label label2;
  }
}