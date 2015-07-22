using OpenHRNexus.Interfaces.Common;

namespace OpenHRNexus.Service.Interfaces {
	public interface INexusDbService {
		INexusUser RequestAccount(string email);
	}
}