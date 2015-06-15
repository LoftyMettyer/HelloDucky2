using System;
using NServiceBus;
using OpenHRNexus.Common.Messaging.Events;

namespace OpenHRNexus.Messaging.Subscriber.Handlers.EventHandlers {
	public class LoginAttemptEventHandler : IHandleMessages<LoginAttemptEvent> {
		public void Handle(LoginAttemptEvent e) {
			Console.WriteLine("Login attempt with outcome " + e.Message);
		}
	}
}
