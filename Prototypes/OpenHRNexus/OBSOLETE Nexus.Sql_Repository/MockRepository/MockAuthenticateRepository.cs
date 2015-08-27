using System;
using System.Collections.Generic;
using Nexus.Repository.Interfaces;
using Nexus.Repository.Messages;
using Repository.Enums;

namespace Nexus.Repository.MockRepository {
	public class MockAuthenticateRepository : IAuthenticateRepository {
		public RegisterNewUserMessage RequestAccount(string email, string userId) {
			return new RegisterNewUserMessage {
				Status = NewUserStatus.Success
			};
		}

		public IEnumerable<string> GetUserPermissions(Guid userId) {
			throw new NotImplementedException();
		}

	}
}