using System.Diagnostics;

namespace OutlookCalendarLogic.Structures {
  public struct ConfigurationError {
	public string Message { get; set; }
	public EventLogEntryType Severity { get; set; }

	public ConfigurationError(string message, EventLogEntryType severity) {
	  this.Message = message;
	  this.Severity = severity;
	}
  }
}
