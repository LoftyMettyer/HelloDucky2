using OpenHRNexus.Interfaces.Common;
using OpenHRNexus.Repository.Interfaces;
using OpenHRNexus.Service.Interfaces;

namespace OpenHRNexus.Service.Services {
	public class AuthenticateService : IAuthenticateService {
		private readonly IAuthenticateRepository _nexusAuthenticateRepository;

		public AuthenticateService(IAuthenticateRepository nexusAuhenticateRepository) {
			_nexusAuthenticateRepository = nexusAuhenticateRepository;
		}

		public INexusUser RequestAccount(string email) {
			return _nexusAuthenticateRepository.RequestAccount(email);
		}
	}
}
