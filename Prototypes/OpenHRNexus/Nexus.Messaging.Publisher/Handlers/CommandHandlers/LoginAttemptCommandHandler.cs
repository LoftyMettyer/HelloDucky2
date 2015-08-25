using System;
using NServiceBus;
using Nexus.Common.Messaging.Commands;
using Nexus.Common.Messaging.Events;

namespace Nexus.Messaging.Publisher.Handlers.CommandHandlers {
	public class LoginAttemptCommandHandler : IHandleMessages<LoginAttemptCommand> {
		public IBus bus;

		public LoginAttemptCommandHandler(IBus bus) {
			this.bus = bus;
		}

		public void Handle(LoginAttemptCommand command) {
			Console.WriteLine("User " + command.UserName + " is trying to login");
			Console.WriteLine("");

			if (command.UserName == "albert" && command.Password == "einstein") {
				bus.Publish(new LoginAttemptEvent() { Message = "SUCCESS" });
			}
			else {
				bus.Publish(new LoginAttemptEvent() { Message = "FAILED" });
			}
		}
	}
}