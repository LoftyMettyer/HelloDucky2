namespace OutlookCalendarService2Tester
{
  partial class MainFormCustomer
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
	  System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MainFormCustomer));
	  this.txtExchangeUserPassword = new System.Windows.Forms.TextBox();
	  this.btnCreateTestCalendarEntry = new System.Windows.Forms.Button();
	  this.btnClearDebugMessages = new System.Windows.Forms.Button();
	  this.label1 = new System.Windows.Forms.Label();
	  this.txtExchangeUser = new System.Windows.Forms.TextBox();
	  this.txtDebugLog = new System.Windows.Forms.TextBox();
	  this.SuspendLayout();
	  // 
	  // txtExchangeUserPassword
	  // 
	  this.txtExchangeUserPassword.Location = new System.Drawing.Point(386, 9);
	  this.txtExchangeUserPassword.Name = "txtExchangeUserPassword";
	  this.txtExchangeUserPassword.PasswordChar = '*';
	  this.txtExchangeUserPassword.Size = new System.Drawing.Size(121, 20);
	  this.txtExchangeUserPassword.TabIndex = 2;
	  // 
	  // btnCreateTestCalendarEntry
	  // 
	  this.btnCreateTestCalendarEntry.Location = new System.Drawing.Point(530, 7);
	  this.btnCreateTestCalendarEntry.Name = "btnCreateTestCalendarEntry";
	  this.btnCreateTestCalendarEntry.Size = new System.Drawing.Size(132, 23);
	  this.btnCreateTestCalendarEntry.TabIndex = 3;
	  this.btnCreateTestCalendarEntry.Text = "Create calendar entry";
	  this.btnCreateTestCalendarEntry.UseVisualStyleBackColor = true;
	  this.btnCreateTestCalendarEntry.Click += new System.EventHandler(this.btnCreateTestCalendarEntry_Click);
	  // 
	  // btnClearDebugMessages
	  // 
	  this.btnClearDebugMessages.Location = new System.Drawing.Point(321, 40);
	  this.btnClearDebugMessages.Name = "btnClearDebugMessages";
	  this.btnClearDebugMessages.Size = new System.Drawing.Size(106, 23);
	  this.btnClearDebugMessages.TabIndex = 4;
	  this.btnClearDebugMessages.Text = "Clear messages";
	  this.btnClearDebugMessages.UseVisualStyleBackColor = true;
	  this.btnClearDebugMessages.Click += new System.EventHandler(this.btnClearDebugMessages_Click);
	  // 
	  // label1
	  // 
	  this.label1.AutoSize = true;
	  this.label1.Location = new System.Drawing.Point(15, 9);
	  this.label1.Name = "label1";
	  this.label1.Size = new System.Drawing.Size(147, 13);
	  this.label1.TabIndex = 15;
	  this.label1.Text = "Exchange user and password";
	  // 
	  // txtExchangeUser
	  // 
	  this.txtExchangeUser.Location = new System.Drawing.Point(173, 9);
	  this.txtExchangeUser.Name = "txtExchangeUser";
	  this.txtExchangeUser.Size = new System.Drawing.Size(207, 20);
	  this.txtExchangeUser.TabIndex = 1;
	  // 
	  // txtDebugLog
	  // 
	  this.txtDebugLog.Font = new System.Drawing.Font("Courier New", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
	  this.txtDebugLog.Location = new System.Drawing.Point(18, 69);
	  this.txtDebugLog.Multiline = true;
	  this.txtDebugLog.Name = "txtDebugLog";
	  this.txtDebugLog.ReadOnly = true;
	  this.txtDebugLog.ScrollBars = System.Windows.Forms.ScrollBars.Both;
	  this.txtDebugLog.Size = new System.Drawing.Size(719, 363);
	  this.txtDebugLog.TabIndex = 5;
	  this.txtDebugLog.WordWrap = false;
	  // 
	  // MainFormCustomer
	  // 
	  this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
	  this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
	  this.ClientSize = new System.Drawing.Size(749, 435);
	  this.Controls.Add(this.txtExchangeUserPassword);
	  this.Controls.Add(this.btnCreateTestCalendarEntry);
	  this.Controls.Add(this.btnClearDebugMessages);
	  this.Controls.Add(this.label1);
	  this.Controls.Add(this.txtExchangeUser);
	  this.Controls.Add(this.txtDebugLog);
	  this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
	  this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
	  this.MaximizeBox = false;
	  this.Name = "MainFormCustomer";
	  this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
	  this.Text = "Outlook Calendar Service Tester";
	  this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.MainFormCustomer_FormClosed);
	  this.ResumeLayout(false);
	  this.PerformLayout();

	}

	#endregion

	private System.Windows.Forms.TextBox txtExchangeUserPassword;
	private System.Windows.Forms.Button btnCreateTestCalendarEntry;
	private System.Windows.Forms.Button btnClearDebugMessages;
	private System.Windows.Forms.Label label1;
	private System.Windows.Forms.TextBox txtExchangeUser;
	private System.Windows.Forms.TextBox txtDebugLog;
  }
}