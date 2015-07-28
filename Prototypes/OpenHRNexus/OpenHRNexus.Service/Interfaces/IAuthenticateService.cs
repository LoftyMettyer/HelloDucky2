using System;
using System.Collections.Generic;
using OpenHRNexus.Repository.Messages;

namespace OpenHRNexus.Service.Interfaces {
	public interface IAuthenticateService {
		RegisterNewUserMessage RequestAccount(string email);
		IEnumerable<string> GetRoles(Guid userId);
	}
}

