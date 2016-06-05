namespace OutlookCalendarService2Tester {
  partial class MainFormSupport {
	/// <summary>
	/// Required designer variable.
	/// </summary>
	private System.ComponentModel.IContainer components = null;

	/// <summary>
	/// Clean up any resources being used.
	/// </summary>
	/// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
	protected override void Dispose(bool disposing) {
	  if (disposing && (components != null)) {
		components.Dispose();
	  }
	  base.Dispose(disposing);
	}

	#region Windows Form Designer generated code

	/// <summary>
	/// Required method for Designer support - do not modify
	/// the contents of this method with the code editor.
	/// </summary>
	private void InitializeComponent() {
         System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MainFormSupport));
         this.btnRunOnce = new System.Windows.Forms.Button();
         this.txtDebugLog = new System.Windows.Forms.TextBox();
         this.txtExchangeUser = new System.Windows.Forms.TextBox();
         this.txtWindowsEventsLog = new System.Windows.Forms.TextBox();
         this.label1 = new System.Windows.Forms.Label();
         this.label2 = new System.Windows.Forms.Label();
         this.label3 = new System.Windows.Forms.Label();
         this.btnTestAutodiscoverLogon = new System.Windows.Forms.Button();
         this.btnClearDebugMessages = new System.Windows.Forms.Button();
         this.btnClearWindowsEventLog = new System.Windows.Forms.Button();
         this.btnCreateTestCalendarEntry = new System.Windows.Forms.Button();
         this.btnCheckConfiguration = new System.Windows.Forms.Button();
         this.txtExchangeUserPassword = new System.Windows.Forms.TextBox();
         this.SuspendLayout();
         // 
         // btnRunOnce
         // 
         this.btnRunOnce.Location = new System.Drawing.Point(548, 22);
         this.btnRunOnce.Name = "btnRunOnce";
         this.btnRunOnce.Size = new System.Drawing.Size(106, 23);
         this.btnRunOnce.TabIndex = 3;
         this.btnRunOnce.Text = "Run once";
         this.btnRunOnce.UseVisualStyleBackColor = true;
         this.btnRunOnce.Click += new System.EventHandler(this.btnRunOnce_Click);
         // 
         // txtDebugLog
         // 
         this.txtDebugLog.Font = new System.Drawing.Font("Courier New", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
         this.txtDebugLog.Location = new System.Drawing.Point(15, 82);
         this.txtDebugLog.Multiline = true;
         this.txtDebugLog.Name = "txtDebugLog";
         this.txtDebugLog.ReadOnly = true;
         this.txtDebugLog.ScrollBars = System.Windows.Forms.ScrollBars.Both;
         this.txtDebugLog.Size = new System.Drawing.Size(525, 363);
         this.txtDebugLog.TabIndex = 9;
         this.txtDebugLog.WordWrap = false;
         // 
         // txtExchangeUser
         // 
         this.txtExchangeUser.Location = new System.Drawing.Point(170, 22);
         this.txtExchangeUser.Name = "txtExchangeUser";
         this.txtExchangeUser.Size = new System.Drawing.Size(207, 20);
         this.txtExchangeUser.TabIndex = 1;
         this.txtExchangeUser.Text = "testsqliis03.cal@company.local";
         // 
         // txtWindowsEventsLog
         // 
         this.txtWindowsEventsLog.Font = new System.Drawing.Font("Courier New", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
         this.txtWindowsEventsLog.Location = new System.Drawing.Point(546, 82);
         this.txtWindowsEventsLog.Multiline = true;
         this.txtWindowsEventsLog.Name = "txtWindowsEventsLog";
         this.txtWindowsEventsLog.ReadOnly = true;
         this.txtWindowsEventsLog.ScrollBars = System.Windows.Forms.ScrollBars.Both;
         this.txtWindowsEventsLog.Size = new System.Drawing.Size(525, 363);
         this.txtWindowsEventsLog.TabIndex = 10;
         this.txtWindowsEventsLog.WordWrap = false;
         // 
         // label1
         // 
         this.label1.AutoSize = true;
         this.label1.Location = new System.Drawing.Point(9, 22);
         this.label1.Name = "label1";
         this.label1.Size = new System.Drawing.Size(147, 13);
         this.label1.TabIndex = 4;
         this.label1.Text = "Exchange user and password";
         // 
         // label2
         // 
         this.label2.AutoSize = true;
         this.label2.Location = new System.Drawing.Point(28, 58);
         this.label2.Name = "label2";
         this.label2.Size = new System.Drawing.Size(89, 13);
         this.label2.TabIndex = 5;
         this.label2.Text = "Debug messages";
         // 
         // label3
         // 
         this.label3.AutoSize = true;
         this.label3.Location = new System.Drawing.Point(545, 58);
         this.label3.Name = "label3";
         this.label3.Size = new System.Drawing.Size(132, 13);
         this.label3.TabIndex = 6;
         this.label3.Text = "Windows Event messages";
         // 
         // btnTestAutodiscoverLogon
         // 
         this.btnTestAutodiscoverLogon.Location = new System.Drawing.Point(660, 22);
         this.btnTestAutodiscoverLogon.Name = "btnTestAutodiscoverLogon";
         this.btnTestAutodiscoverLogon.Size = new System.Drawing.Size(132, 23);
         this.btnTestAutodiscoverLogon.TabIndex = 4;
         this.btnTestAutodiscoverLogon.Text = "Test autodiscover logon";
         this.btnTestAutodiscoverLogon.UseVisualStyleBackColor = true;
         this.btnTestAutodiscoverLogon.Click += new System.EventHandler(this.btnTestAutodiscoverLogon_Click);
         // 
         // btnClearDebugMessages
         // 
         this.btnClearDebugMessages.Location = new System.Drawing.Point(123, 53);
         this.btnClearDebugMessages.Name = "btnClearDebugMessages";
         this.btnClearDebugMessages.Size = new System.Drawing.Size(106, 23);
         this.btnClearDebugMessages.TabIndex = 7;
         this.btnClearDebugMessages.Text = "Clear messages";
         this.btnClearDebugMessages.UseVisualStyleBackColor = true;
         this.btnClearDebugMessages.Click += new System.EventHandler(this.btnClearDebugMessages_Click);
         // 
         // btnClearWindowsEventLog
         // 
         this.btnClearWindowsEventLog.Location = new System.Drawing.Point(678, 53);
         this.btnClearWindowsEventLog.Name = "btnClearWindowsEventLog";
         this.btnClearWindowsEventLog.Size = new System.Drawing.Size(106, 23);
         this.btnClearWindowsEventLog.TabIndex = 8;
         this.btnClearWindowsEventLog.Text = "Clear messages";
         this.btnClearWindowsEventLog.UseVisualStyleBackColor = true;
         this.btnClearWindowsEventLog.Click += new System.EventHandler(this.btnClearWindowsEventLog_Click);
         // 
         // btnCreateTestCalendarEntry
         // 
         this.btnCreateTestCalendarEntry.Location = new System.Drawing.Point(936, 22);
         this.btnCreateTestCalendarEntry.Name = "btnCreateTestCalendarEntry";
         this.btnCreateTestCalendarEntry.Size = new System.Drawing.Size(132, 23);
         this.btnCreateTestCalendarEntry.TabIndex = 6;
         this.btnCreateTestCalendarEntry.Text = "Create calendar entry";
         this.btnCreateTestCalendarEntry.UseVisualStyleBackColor = true;
         this.btnCreateTestCalendarEntry.Click += new System.EventHandler(this.btnCreateTestCalendarEntry_Click);
         // 
         // btnCheckConfiguration
         // 
         this.btnCheckConfiguration.Location = new System.Drawing.Point(798, 22);
         this.btnCheckConfiguration.Name = "btnCheckConfiguration";
         this.btnCheckConfiguration.Size = new System.Drawing.Size(132, 23);
         this.btnCheckConfiguration.TabIndex = 5;
         this.btnCheckConfiguration.Text = "Check configuration";
         this.btnCheckConfiguration.UseVisualStyleBackColor = true;
         this.btnCheckConfiguration.Click += new System.EventHandler(this.btnCheckConfiguration_Click);
         // 
         // txtExchangeUserPassword
         // 
         this.txtExchangeUserPassword.Location = new System.Drawing.Point(383, 22);
         this.txtExchangeUserPassword.Name = "txtExchangeUserPassword";
         this.txtExchangeUserPassword.PasswordChar = '*';
         this.txtExchangeUserPassword.Size = new System.Drawing.Size(121, 20);
         this.txtExchangeUserPassword.TabIndex = 2;
         this.txtExchangeUserPassword.Text = "Connect99";
         // 
         // MainFormSupport
         // 
         this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
         this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
         this.ClientSize = new System.Drawing.Size(1083, 455);
         this.Controls.Add(this.txtExchangeUserPassword);
         this.Controls.Add(this.btnCheckConfiguration);
         this.Controls.Add(this.btnCreateTestCalendarEntry);
         this.Controls.Add(this.btnClearWindowsEventLog);
         this.Controls.Add(this.btnClearDebugMessages);
         this.Controls.Add(this.btnTestAutodiscoverLogon);
         this.Controls.Add(this.label3);
         this.Controls.Add(this.label2);
         this.Controls.Add(this.label1);
         this.Controls.Add(this.txtWindowsEventsLog);
         this.Controls.Add(this.txtExchangeUser);
         this.Controls.Add(this.txtDebugLog);
         this.Controls.Add(this.btnRunOnce);
         this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
         this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
         this.MaximizeBox = false;
         this.Name = "MainFormSupport";
         this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
         this.Text = "Outlook Calendar Service Tester";
         this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.MainFormSupport_FormClosed);
         this.ResumeLayout(false);
         this.PerformLayout();

	}

	#endregion

	private System.Windows.Forms.Button btnRunOnce;
	private System.Windows.Forms.TextBox txtDebugLog;
	private System.Windows.Forms.TextBox txtExchangeUser;
	private System.Windows.Forms.TextBox txtWindowsEventsLog;
	private System.Windows.Forms.Label label1;
	private System.Windows.Forms.Label label2;
	private System.Windows.Forms.Label label3;
	private System.Windows.Forms.Button btnTestAutodiscoverLogon;
	private System.Windows.Forms.Button btnClearDebugMessages;
	private System.Windows.Forms.Button btnClearWindowsEventLog;
	private System.Windows.Forms.Button btnCreateTestCalendarEntry;
	private System.Windows.Forms.Button btnCheckConfiguration;
	private System.Windows.Forms.TextBox txtExchangeUserPassword;
  }
}

