using OpenHRNexus.Interfaces.Common;

namespace OpenHRNexus.Repository.Interfaces {
	public interface IAuthenticateRepository {
		INexusUser RequestAccount(string email);
	}
}
