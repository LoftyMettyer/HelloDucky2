using System;
using System.Collections.Generic;
using OpenHRNexus.Repository.Messages;

namespace OpenHRNexus.Service.Interfaces {
	public interface IAuthenticateService {
		RegisterNewUserMessage RequestAccount(string email, string userId);
		IEnumerable<string> GetClaims(Guid userId);
	}
}

