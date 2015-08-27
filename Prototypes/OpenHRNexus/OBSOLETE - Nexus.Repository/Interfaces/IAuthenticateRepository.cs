using System;
using System.Collections.Generic;
using Nexus.Repository.Messages;

namespace Nexus.Repository.Interfaces {
	public interface IAuthenticateRepository {
		RegisterNewUserMessage RequestAccount(string email, string userId);
		IEnumerable<string> GetUserPermissions(Guid userId);
	}
}
