using NServiceBus;

namespace OpenHRNexus.Common.Messaging.Commands {
	public class LoginAttemptCommand : ICommand {
		public string UserName;
		public string Password;
	}
}
