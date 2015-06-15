using NServiceBus;

namespace OpenHRNexus.Common.Messaging.Events {
	public class LoginAttemptEvent : IEvent {
		public string Message;
	}
}
