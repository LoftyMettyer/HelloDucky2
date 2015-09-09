﻿using System;
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
			var userID = new Guid("088C6A78-E14A-41B0-AD93-4FB7D3ADE96C");

			var welcomeMessage = nexusDb.GetWelcomeMessageData(userID, "EN-GB");
			Assert.IsNotNull(welcomeMessage);

		}

		[TestMethod]
		[Description("Gets roles from the repository")]
		public void RepositoryGetPermissions() {

			var nexusDb = new SqlAuthenticateRepository();
			var userID = new Guid("088C6A78-E14A-41B0-AD93-4FB7D3ADE96C");

			var roles = nexusDb.GetUserPermissions(userID);
			Assert.IsNotNull(roles);

		}
	}
}
