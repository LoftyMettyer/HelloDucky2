using Microsoft.VisualStudio.TestTools.UnitTesting;
using Nexus.Common.Enums;

namespace Nexus.Sql_Repository.Tests.Users {
	[TestClass]
	public class RegisterUser {

		[TestMethod]
		public void WelcomeRepository_RegisterUnknownUser() {
			var actualDb = new SqlWelcomeRepository();
			var message = actualDb.RequestAccount("nosuchuserexists@notsuchcompany.com", new System.Guid().ToString());
			Assert.AreEqual(message.Status, NewUserStatus.UnrecognizedEmail);

		}

		[TestMethod]
		public void WelcomeRepository_RegisterExistingUser() {
			var newUser = new SqlWelcomeRepository();
			var message = newUser.RequestAccount("Nick.Gibson@advancedcomputersoftware.com", new System.Guid().ToString());
			Assert.AreEqual(message.Status, NewUserStatus.AlreadyExists);
		}

		[TestMethod]
		public void WelcomeRepository_RegisterValidNewUser() {
			var newUser = new SqlWelcomeRepository();
			var message = newUser.RequestAccount("Alexandre.Abley@HelloDuckyWorld.com", new System.Guid().ToString());
			Assert.AreEqual(message.Status, NewUserStatus.Success);
		}

	}

}
