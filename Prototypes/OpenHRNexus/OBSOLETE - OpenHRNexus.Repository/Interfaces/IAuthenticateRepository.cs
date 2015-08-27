using System;
using System.Collections.Generic;
using OpenHRNexus.Repository.Messages;

namespace OpenHRNexus.Repository.Interfaces {
	public interface IAuthenticateRepository {
		RegisterNewUserMessage RequestAccount(string email, string userId);
		IEnumerable<string> GetUserPermissions(Guid userId);
	}
}
