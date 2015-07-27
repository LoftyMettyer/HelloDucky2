using OpenHRNexus.Repository.Messages;

namespace OpenHRNexus.Service.Interfaces {
	public interface IAuthenticateService {
		RegisterNewUserMessage RequestAccount(string email);
	}
}

