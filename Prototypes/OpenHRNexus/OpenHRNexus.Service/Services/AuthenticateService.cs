using OpenHRNexus.Interfaces.Common;
using OpenHRNexus.Repository.Interfaces;
using OpenHRNexus.Service.Interfaces;

namespace OpenHRNexus.Service.Services {
	public class AuthenticateService : IAuthenticateService {
		private readonly IAuthenticateRepository _authenticateRepository;

		public AuthenticateService(IAuthenticateRepository auhenticateRepository) {
			_authenticateRepository = auhenticateRepository;
		}

		public INexusUser RequestAccount(string email) {
			return _authenticateRepository.RequestAccount(email);
		}
	}
}
