using System;
using System.Collections.Generic;
using Nexus.Common.Messages;

namespace Nexus.Common.Interfaces.Services
{
	public interface IWelcomeService {
		RegisterNewUserMessage RequestAccount(string email, string userId);
		IEnumerable<string> GetClaims(Guid userId);
	}
}

