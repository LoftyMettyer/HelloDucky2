using System;
using System.Collections.Generic;
using OpenHRNexus.Repository.Interfaces;
using OpenHRNexus.Repository.Messages;
using OpenHRNexus.Service.Interfaces;

namespace OpenHRNexus.Service.Services {
	public class AuthenticateService : IAuthenticateService {
		private readonly IAuthenticateRepository _authenticateRepository;

		public AuthenticateService(IAuthenticateRepository auhenticateRepository) {
			_authenticateRepository = auhenticateRepository;
		}

		public RegisterNewUserMessage RequestAccount(string email) {
			return _authenticateRepository.RequestAccount(email);
		}

		public IEnumerable<string> GetRoles(Guid userId)
		{
			return _authenticateRepository.GetUserPermissions(userId);
		}
	}
}
