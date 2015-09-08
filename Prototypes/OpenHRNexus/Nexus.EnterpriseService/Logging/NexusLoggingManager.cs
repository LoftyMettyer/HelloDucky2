using System.Diagnostics;
using Microsoft.Practices.EnterpriseLibrary.Logging;
using Microsoft.Practices.EnterpriseLibrary.Logging.Formatters;
using Microsoft.Practices.EnterpriseLibrary.Logging.TraceListeners;

namespace Nexus.EnterpriseService.Logging {
	public class NexusLoggingManager {
		public LogWriter LogWriter;

		public NexusLoggingManager(string subSystem) {
			//To differentiate between Event Log Sources add a suffix to the SubSystem
			subSystem += ".Logging";

			LoggingConfiguration loggingConfiguration = BuildLoggingConfig(subSystem);
			LogWriter = new LogWriter(loggingConfiguration);
		}

		private static LoggingConfiguration BuildLoggingConfig(string subSystem) {
			//Create a new event log and source ("subsystem") if it doesn't exist
			//if (!EventLog.Exists(NexusEnterpriseConstants.WindowsEventLogName)) {
			//	if (!EventLog.SourceExists(subSystem)) {
			//		EventLog.CreateEventSource(subSystem, NexusEnterpriseConstants.WindowsEventLogName);
			//	}
			//}

			// Formatters
			TextFormatter briefFormatter = new TextFormatter("Timestamp: {timestamp(local)}{newline}Message: {message}{newline}Category: {category}{newline}Priority: {priority}{newline}EventId: {eventid}{newline}ActivityId: {property(ActivityId)}{newline}Severity: {severity}{newline}Title:{title}{newline}");
			TextFormatter extendedFormatter = new TextFormatter("Timestamp: {timestamp}{newline}Message: {message}{newline}Category: {category}{newline}Priority: {priority}{newline}EventId: {eventid}{newline}Severity: {severity}{newline}Title: {title}{newline}Activity ID: {property(ActivityId)}{newline}Machine: {localMachine}{newline}App Domain: {localAppDomain}{newline}ProcessId: {localProcessId}{newline}Process Name: {localProcessName}{newline}Thread Name: {threadName}{newline}Win32 ThreadId:{win32ThreadId}{newline}Extended Properties: {dictionary({key} - {value}{newline})}");

			// Trace Listeners
			// var flatFileTraceListener = new FlatFileTraceListener(@"Nexus.log", "----------------------------------------", "----------------------------------------", briefFormatter);
			var eventLog = new EventLog {
				Log = NexusEnterpriseConstants.WindowsEventLogName,
				Source = subSystem
			};
			var eventLogTraceListener = new FormattedEventLogTraceListener(eventLog);

			// Build Configuration
			var loggingConfig = new LoggingConfiguration();
			loggingConfig.AddLogSource(NexusEnterpriseConstants.WindowsEventLogName, SourceLevels.All, false).AddTraceListener(eventLogTraceListener);
			// loggingConfig.AddLogSource("DiskFiles", SourceLevels.All, true).AddTraceListener(flatFileTraceListener);

			return loggingConfig;
		}
	}
}
