using OpenHRNexus.Interfaces.Common;
using OpenHRNexus.Service.Interfaces;

namespace OpenHRNexus.Service.Services {
	public class AuthenticateService : INexusDbService {
		private readonly INexusDbService nexusDbService;

		public AuthenticateService(INexusDbService nexusDbService) {
			nexusDbService = nexusDbService;
		}

		public INexusUser RequestAccount(string email) {
			return nexusDbService.RequestAccount(email);
		}
	}
}
