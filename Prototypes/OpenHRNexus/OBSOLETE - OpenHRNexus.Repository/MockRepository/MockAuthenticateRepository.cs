using System;
using System.Collections.Generic;
using OpenHRNexus.Repository.Interfaces;
using OpenHRNexus.Repository.Messages;
using Repository.Enums;

namespace OpenHRNexus.Repository.MockRepository {
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