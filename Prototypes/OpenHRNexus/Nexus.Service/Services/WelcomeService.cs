using System;
using System.Collections.Generic;
using Nexus.Common.Interfaces.Repository;
using Nexus.Common.Messages;
using Nexus.Common.Interfaces.Services;

namespace Nexus.Service.Services {
	public class WelcomeService : IWelcomeService {
		private readonly IWelcomeRepository _welcomeRepository;

		public WelcomeService(IWelcomeRepository welcomeRepository) {
            _welcomeRepository = welcomeRepository;
		}

		public RegisterNewUserMessage RequestAccount(string email, string userId) {
			return _welcomeRepository.RequestAccount(email, userId);
		}

		public IEnumerable<string> GetClaims(Guid userId) {
			return _welcomeRepository.GetUserPermissions(userId);
		}
	}
}
