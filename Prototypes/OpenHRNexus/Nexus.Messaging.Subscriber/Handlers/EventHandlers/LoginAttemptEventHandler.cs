using System;
using NServiceBus;
using Nexus.Common.Messaging.Events;

namespace Nexus.Messaging.Subscriber.Handlers.EventHandlers {
	public class LoginAttemptEventHandler : IHandleMessages<LoginAttemptEvent> {
		public void Handle(LoginAttemptEvent e) {
			Console.WriteLine("Login attempt with outcome " + e.Message);
		}
	}
}
