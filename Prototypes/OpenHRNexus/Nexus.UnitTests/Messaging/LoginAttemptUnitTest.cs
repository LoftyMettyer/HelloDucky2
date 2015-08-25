using Microsoft.VisualStudio.TestTools.UnitTesting;
using NServiceBus;
using NServiceBus.Testing;
using Nexus.Common.Messaging.Commands;
using Nexus.Common.Messaging.Events;
using Nexus.Messaging.Subscriber.Handlers.EventHandlers;

namespace Nexus.UnitTests.Messaging {
	[TestClass]
	public class LoginAttemptUnitTest {
		[TestMethod]
		public void LoginAttempt() {

			//Test.Initialize();
			//Test.Handler<LoginAttemptEventHandler>().ExpectReply<LoginAttemptEvent>(e => e.Message == "SUCCESS");
			//var loginAttempt = new LoginAttemptCommand() { UserName = "peter", Password = "pan" }; //This user will succeed
			//var bus =      Bus.Create
			//	bus.Send("Nexus.Messaging.Publisher", loginAttempt);

			//Assert
		}
	}
}
