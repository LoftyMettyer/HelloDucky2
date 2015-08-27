using Microsoft.VisualStudio.TestTools.UnitTesting;
using NServiceBus;
using NServiceBus.Testing;
using OpenHRNexus.Common.Messaging.Commands;
using OpenHRNexus.Common.Messaging.Events;
using OpenHRNexus.Messaging.Subscriber.Handlers.EventHandlers;

namespace OpenHRNexus.UnitTests.Messaging {
	[TestClass]
	public class LoginAttemptUnitTest {
		[TestMethod]
		public void LoginAttempt() {

			//Test.Initialize();
			//Test.Handler<LoginAttemptEventHandler>().ExpectReply<LoginAttemptEvent>(e => e.Message == "SUCCESS");
			//var loginAttempt = new LoginAttemptCommand() { UserName = "peter", Password = "pan" }; //This user will succeed
			//var bus =      Bus.Create
			//	bus.Send("OpenHRNexus.Messaging.Publisher", loginAttempt);

			//Assert
		}
	}
}
