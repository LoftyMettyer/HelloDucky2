using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Nexus.Sql_Repository.Tests.Users {
	[TestClass]
	public class ConnectUser {

		[TestMethod]
		public void CanCreateSqlAuthenticateRepository() {
			var nexusDb = new SqlAuthenticateRepository();
			Assert.IsNotNull(nexusDb);
		}


		[TestMethod]
		public void ConnectAsValidUser_English() {

			var nexusDb = new SqlAuthenticateRepository();
			var userID = new Guid("E206BE61-3591-4A7D-B2A8-F5004CC0A7ED");

			var welcomeMessage = nexusDb.GetWelcomeMessageData(userID, "EN-GB");
			Assert.IsNotNull(welcomeMessage);

		}

		[TestMethod]
		[Description("Gets roles from the repository")]
		public void RepositoryGetPermissions() {

			var nexusDb = new SqlAuthenticateRepository();
			var userID = new Guid("3CEF8BE0-B512-4A4E-8E02-F1C0E31BBB30");

			var roles = nexusDb.GetUserPermissions(userID);
			Assert.IsNotNull(roles);

		}
	}
}
