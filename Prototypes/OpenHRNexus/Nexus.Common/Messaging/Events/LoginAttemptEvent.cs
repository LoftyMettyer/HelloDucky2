using NServiceBus;

namespace Nexus.Common.Messaging.Events {
	public class LoginAttemptEvent : IEvent {
		public string Message;
	}
}
