using System;
using OpenHRNexus.Interfaces.Common;
using OpenHRNexus.Repository.DatabaseClasses;
using OpenHRNexus.Repository.Interfaces;

namespace OpenHRNexus.Repository.MockRepository {
	public class MockAuthenticateRepository : IAuthenticateRepository {
		public INexusUser RequestAccount(string email) {
			return new User {
				Id = Guid.NewGuid(),
				Role = "Employee",
				LastConnectDateTime = DateTime.Now
			};
		}
	}
}