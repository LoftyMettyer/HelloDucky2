using Microsoft.VisualStudio.TestTools.UnitTesting;
using Nexus.Repository.Interfaces;
using Moq;
using Nexus.Repository.SQLServer;
using Repository.Enums;

namespace Nexus.Repository.Tests.Users {
	[TestClass]
	public class RegisterUser {
		//[TestMethod]
		//public void ConnectToNexusUserDb()
		//{
		//	var nexusDb = new NexusUserEntities();
		//	Assert.IsNotNull(nexusDb);
		//}

		[TestMethod]
		public void RegisterUnknownUser() {
			var actualDb = new SqlAuthenticateRepository();
			var message = actualDb.RequestAccount("nosuchuserexists@notsuchcompany.com", new System.Guid().ToString());
			Assert.AreEqual(message.Status, NewUserStatus.UnrecognizedEmail);

			//		var mockDb = new Mock<IAuthenticateRepository>().Object;
			//		var newUserMessage = mockDb.RequestAccount("nosuchuserexists@notsuchcompany.com");
			//		Assert.AreEqual(newUserMessage.Status, NewUserStatus.UnrecognizedEmail);
		}

		[TestMethod]
		public void RegisterExistingUser() {
			var newUser = new SqlAuthenticateRepository();
			var message = newUser.RequestAccount("John.Adams@HelloDuckyWorld.com", new System.Guid().ToString());
			Assert.AreEqual(message.Status, NewUserStatus.AlreadyExists);
		}

		[TestMethod]
		public void RegisterValidNewUser() {
			var newUser = new SqlAuthenticateRepository();
			var message = newUser.RequestAccount("Alexandre.Abley@HelloDuckyWorld.com", new System.Guid().ToString());
			Assert.AreEqual(message.Status, NewUserStatus.Success);
		}

	}

}
