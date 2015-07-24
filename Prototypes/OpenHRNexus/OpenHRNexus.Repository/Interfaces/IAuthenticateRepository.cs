using OpenHRNexus.Interfaces.Common;

//Just a comment
namespace OpenHRNexus.Repository.Interfaces {
	public interface IAuthenticateRepository {
		INexusUser RequestAccount(string email);
	}
}
