using OpenHRNexus.Interfaces.Common;

namespace OpenHRNexus.Service.Interfaces {
	public interface IAuthenticateService {
		INexusUser RequestAccount(string email);
	}
}