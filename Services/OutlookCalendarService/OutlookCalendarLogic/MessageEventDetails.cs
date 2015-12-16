using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OutlookCalendarLogic {
  public class MessageEventDetails {
	public enum MessageEventType {
	  WindowsEventsLog,
	  DebugLog,
	  WindowsEventsLogAndDebugLog
	}

	public string Message { get; set; }
	public EventLogEntryType Severity { get; set; }
	public MessageEventType TriggeredEventType { get; set; }

	/// <summary>
	/// Create a new instance of the MessageEventDetails
	/// </summary>
	/// <param name="message"></param>
	/// <param name="severity"></param>
	/// <param name="triggeredEventType"></param>
	public MessageEventDetails(string message, EventLogEntryType severity, MessageEventType triggeredEventType) {
	  Message = message;
	  Severity = severity;
	  TriggeredEventType = triggeredEventType;
	}

	/// <summary>
	/// Create a new instance of the MessageEventDetails with a default severity of 'Information'
	/// </summary>
	/// <param name="message"></param>
	/// <param name="triggeredEventType"></param>
	/// <param name="severity"></param>
	public MessageEventDetails(string message, MessageEventType triggeredEventType, EventLogEntryType severity = EventLogEntryType.Information) {
	  Message = message;
	  Severity = severity;
	  TriggeredEventType = triggeredEventType;
	}

	/// <summary>
	/// Create a new instance of the MessageEventDetails with a default severity of 'Information' and a triggered even of 'DebugLog'
	/// </summary>
	/// <param name="message"></param>
	public MessageEventDetails(string message) {
	  Message = message;
	  Severity = EventLogEntryType.Information;
	  TriggeredEventType = MessageEventType.DebugLog;
	}
  }
}
