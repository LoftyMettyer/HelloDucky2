namespace OutlookCalendarService2 {
  partial class CalendarService {
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

	#region Component Designer generated code

	/// <summary> 
	/// Required method for Designer support - do not modify 
	/// the contents of this method with the code editor.
	/// </summary>
	private void InitializeComponent() {
	  this.SrvTmr = new System.Timers.Timer();
	  ((System.ComponentModel.ISupportInitialize)(this.SrvTmr)).BeginInit();
	  // 
	  // SrvTmr
	  // 
	  this.SrvTmr.Enabled = false;
	  this.SrvTmr.Elapsed += new System.Timers.ElapsedEventHandler(this.SrvTmr_Elapsed);
	  // 
	  // CalendarService
	  // 
	  this.ServiceName = "OpenHR Outlook Calendar Service 2";
	  ((System.ComponentModel.ISupportInitialize)(this.SrvTmr)).EndInit();

	}

	#endregion

	private System.Timers.Timer SrvTmr;
  }
}
