using System;
using System.Collections.Generic;
using Nexus.Repository.Messages;

namespace Nexus.Service.Interfaces {
	public interface IAuthenticateService {
		RegisterNewUserMessage RequestAccount(string email, string userId);
		IEnumerable<string> GetClaims(Guid userId);
	}
}

