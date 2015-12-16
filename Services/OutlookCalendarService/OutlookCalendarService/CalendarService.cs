using System;
using System.Diagnostics;
using System.IO;
using System.Security.Principal;
using System.ServiceProcess;
using System.DirectoryServices.AccountManagement;
using OutlookCalendarLogic;
using OutlookCalendarLogic.Structures;

namespace OutlookCalendarService2 {
  public partial class CalendarService : ServiceBase {
	private Worker _worker;
	private StreamWriter _traceWriter;
	private const int PollingInterval = 60000;

	public CalendarService() {
	  InitializeComponent();
	}

	protected override void OnStart(string[] args) {
	  //Uncomment line below while debugging so when the service starts you have about 13 seconds to attach the VS debugger
	  //System.Threading.Thread.Sleep(13000); //////////////////////////////////////////////////////////////////
	  
	  //First, set current directory to be the same directory where the service is running; by default, the current directory for a Windows service is the System32 folder
	  Directory.SetCurrentDirectory(AppDomain.CurrentDomain.BaseDirectory);

	  //Second, check that we have permission to create diagnostics files in the current folder
	  try {
		var f = File.Create(Directory.GetCurrentDirectory() + "\\Temp.txt");
		f.Close();
		f.Dispose();
		File.Delete(Directory.GetCurrentDirectory() + "\\Temp.txt");
	  } catch (Exception ex) {
		LogMessageToWindowsEventLog("Service couldn't start with error message \"" + ex.Message + "\"", EventLogEntryType.Error);
		this.Stop();
		return;
	  }

	  string s = "";

	  try {
		//Get the email address of the account under which the service is running; we need it to autodiscover Exchange
		//WindowsIdentity.GetCurrent() returns the user in the format domain\user
		string[] domainAndUserNameArray = WindowsIdentity.GetCurrent().Name.Split('\\');
		PrincipalContext dc = new PrincipalContext(ContextType.Domain);
		UserPrincipal user = UserPrincipal.FindByIdentity(dc, domainAndUserNameArray[1]);
		string userEmail = user.EmailAddress;

		_worker = new Worker(userEmail, false);
		_worker.RaiseMessageEvent += HandleRaisedMessageEvent;

		if (!_worker.GetAndCheckUserConfiguration()) {
		  foreach (ConfigurationError err in _worker.ConfigurationErrors) {
			s += err.Message + Environment.NewLine;
		  }

		  LogMessageToDebugFile("Service stopping due to the following configuration errors:" + Environment.NewLine + s, EventLogEntryType.Error);
		  LogMessageToWindowsEventLog("Service stopping due to the following configuration errors:" + Environment.NewLine + s, EventLogEntryType.Error);

		  //Stop the service
		  this.Stop();
		}

		if (_worker.UserDefinedConfiguration.Debug)
		  _traceWriter = File.AppendText(Directory.GetCurrentDirectory() + "\\OutlookSvc_Debug.txt");

		LogMessageToDebugFile($"OpenHR Outlook Calendar Service 2 ({_worker.ServiceVersion}) started successfully.", EventLogEntryType.Information);
		LogMessageToWindowsEventLog($"OpenHR Outlook Calendar Service 2 ({_worker.ServiceVersion}) started successfully.", EventLogEntryType.Information);

		_worker.OutputUserConfiguration();
		_worker.DescribeServicedSystems();

		SrvTmr.Interval = PollingInterval;
		SrvTmr.Enabled = true;

		//_worker.ProcessEntries();
		//_worker.Stop();
		//this.Stop();
	  } catch (Exception ex) {
		LogMessageToWindowsEventLog("Service couldn't start with error message \"" + ex.Message + "\"", EventLogEntryType.Error);
		this.Stop();
	  }
	}

	private void HandleRaisedMessageEvent(MessageEventDetails e) {
	  if (_worker.UserDefinedConfiguration.Debug && (e.TriggeredEventType == MessageEventDetails.MessageEventType.DebugLog || e.TriggeredEventType == MessageEventDetails.MessageEventType.WindowsEventsLogAndDebugLog)) {
		LogMessageToDebugFile(e.Message, e.Severity);
	  }

	  if (e.TriggeredEventType == MessageEventDetails.MessageEventType.WindowsEventsLog || e.TriggeredEventType == MessageEventDetails.MessageEventType.WindowsEventsLogAndDebugLog) {
		LogMessageToWindowsEventLog(e.Message, e.Severity);
	  }

	  if (e.Severity == EventLogEntryType.Error) //Stop service for critical errors
	  {
		LogMessageToDebugFile(e.Message + Environment.NewLine + "Critical error found, service stopping", EventLogEntryType.Error);
		LogMessageToWindowsEventLog(e.Message + Environment.NewLine + "Critical error found, service stopping", EventLogEntryType.Error);
		this.Stop();
	  }
	}

	private void LogMessageToDebugFile(string message, EventLogEntryType severity, bool addDateTimeStamp = true) {
	  if (!_worker.UserDefinedConfiguration.Debug) //Only output if trace is enabled
		return;

	  if (severity == EventLogEntryType.Information) {
		if (addDateTimeStamp)
		  _traceWriter.WriteLine("{0} {1}: {2}", DateTime.Now.ToShortDateString(), DateTime.Now.ToLongTimeString(), message);
		else
		  _traceWriter.WriteLine("{0}", message);
	  } else {
		if (addDateTimeStamp)
		  _traceWriter.WriteLine("{0} {1}: [{2}] {3}", DateTime.Now.ToShortDateString(), DateTime.Now.ToLongTimeString(), severity, message);
		else
		  _traceWriter.WriteLine("[{0}] {1}", severity, message);
	  }

	  //Update the underlying file.
	  _traceWriter.Flush();
	}

	private void LogMessageToWindowsEventLog(string message, EventLogEntryType severity) {
	  EventLog log = new EventLog();
	  string logName = "Advanced Business Solutions";
	  string source = "OpenHR Outlook Calendar Service";

	  log.Log = logName;
	  log.Source = source;
	  log.WriteEntry(message, severity);
	  log.Close();
	}

	protected override void OnStop() {
	  _worker.Stop();

	  if (_worker.UserDefinedConfiguration.Debug)
		_traceWriter?.Close();
	}

	private void SrvTmr_Elapsed(object sender, System.Timers.ElapsedEventArgs e) {
	  _worker.ProcessEntries();
	}
  }
}
