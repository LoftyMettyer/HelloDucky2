using System;
using System.Collections.Generic;
using Nexus.Common.Messages;

namespace Nexus.Common.Interfaces.Repository {
	public interface IAuthenticateRepository {
		RegisterNewUserMessage RequestAccount(string email, string userId);
		IEnumerable<string> GetUserPermissions(Guid userId);
	}
}
