﻿using System;
using System.Collections.Generic;
using System.Diagnostics;
using Microsoft.Practices.EnterpriseLibrary.ExceptionHandling;
using Microsoft.Practices.EnterpriseLibrary.ExceptionHandling.Logging;
using Microsoft.Practices.EnterpriseLibrary.Logging;
using Microsoft.Practices.EnterpriseLibrary.Logging.TraceListeners;

namespace OpenHRNexus.EnterpriseServices.ExceptionHandling {
	public class NexusExceptionManager {
		public ExceptionManager ExceptionManager;

		/// <summary>
		/// Create a new instance of the NexusExceptionManager
		/// </summary>
		/// <param name="subSystem">The name of the OpenHRNexus SubSystem to log error information for (Repository, WebAPI, etc.)</param>
		/// <param name="useEventLog">Use the event log?</param>
		/// <param name="showExceptionPolicyName">If set to true, also log the ExceptionPolicyName (useful for debugging). Default is false</param>
		public NexusExceptionManager(string subSystem, bool useEventLog = false, bool showExceptionPolicyName = false) {
			LoggingConfiguration loggingConfiguration = BuildLoggingConfig(subSystem);
			LogWriter logWriter = new LogWriter(loggingConfiguration);

			// Create the default ExceptionManager object programatically
			ExceptionManager = BuildExceptionManagerConfig(logWriter, subSystem, useEventLog, showExceptionPolicyName);
		}

		private static LoggingConfiguration BuildLoggingConfig(string subSystem) {
			//Create a new event log and source ("subsystem") if it doesn't exist
			if (!EventLog.SourceExists(subSystem)) {
				EventLog.CreateEventSource(subSystem, NexusExceptionHandlingConstants.WindowsEventLogName);
			}

			var eventLog = new EventLog {
				Log = NexusExceptionHandlingConstants.WindowsEventLogName,
				Source = subSystem
			};
			var eventLogTraceListener = new FormattedEventLogTraceListener(eventLog);

			//Build Configuration
			var loggingConfig = new LoggingConfiguration();
			loggingConfig.AddLogSource(NexusExceptionHandlingConstants.WindowsEventLogName, SourceLevels.All, false).AddTraceListener(eventLogTraceListener);

			//Special Sources Configuration
			loggingConfig.SpecialSources.LoggingErrorsAndWarnings.AddTraceListener(eventLogTraceListener);

			return loggingConfig;
		}

		private static ExceptionManager BuildExceptionManagerConfig(LogWriter logWriter, string subSystem, bool useEventLog, bool showExceptionPolicyName) {
			var policies = new List<ExceptionPolicyDefinition>();

			//Logging Exception Handler: to be added to the different policies if required
			var loggingExceptionHandler = new LoggingExceptionHandler(NexusExceptionHandlingConstants.WindowsEventLogName, 9001, TraceEventType.Error, subSystem, 5, typeof(TextExceptionFormatter), logWriter);

			//AssistingAdministrators policy
			var assistingAdministratorsExceptionHandlers = new List<IExceptionHandler>();
			if (useEventLog) {
				assistingAdministratorsExceptionHandlers.Add(loggingExceptionHandler);
			}
			assistingAdministratorsExceptionHandlers.Add(
				new ReplaceHandler(
					(showExceptionPolicyName ? "[AssistingAdministrators] " : "") +
					"Application error. Please advise your administrator and provide them with this error code: {handlingInstanceID}",
					typeof(Exception)));
			var assistingAdministrators = new List<ExceptionPolicyEntry> {
				new ExceptionPolicyEntry(typeof(Exception),
					PostHandlingAction.ThrowNewException,
					assistingAdministratorsExceptionHandlers)
			};
			policies.Add(new ExceptionPolicyDefinition(NexusExceptionPolicyNames.AssistingAdministrators.ToString(), assistingAdministrators));

			//ExceptionShielding policy
			var exceptionShieldingExceptionHandlers = new List<IExceptionHandler>();
			if (useEventLog) {
				exceptionShieldingExceptionHandlers.Add(loggingExceptionHandler);
			}
			exceptionShieldingExceptionHandlers.Add(
				new WrapHandler(
					(showExceptionPolicyName ? "[ExceptionShielding] " : "") +
					"Application error. Please contact your administrator.",
					typeof(Exception)));
			var exceptionShielding = new List<ExceptionPolicyEntry>{
				new ExceptionPolicyEntry(typeof(Exception),
					PostHandlingAction.ThrowNewException,
					exceptionShieldingExceptionHandlers)
			};
			policies.Add(new ExceptionPolicyDefinition(NexusExceptionPolicyNames.ExceptionShielding.ToString(), exceptionShielding));

			//
			return new ExceptionManager(policies);
		}
	}
}
