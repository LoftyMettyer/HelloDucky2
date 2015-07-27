using System;
using OpenHRNexus.Repository.Messages;

namespace OpenHRNexus.Repository.Interfaces {
	public interface IAuthenticateRepository {
		RegisterNewUserMessage RequestAccount(string email);
	}
}
