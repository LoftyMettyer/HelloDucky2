using System;
using System.Collections.Generic;
using Nexus.Common.Enums;
using Nexus.Common.Interfaces.Repository;
using Nexus.Common.Messages;

namespace Nexus.Mock_Repository {
	public class MockWelcomeRepository : IWelcomeRepository {
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