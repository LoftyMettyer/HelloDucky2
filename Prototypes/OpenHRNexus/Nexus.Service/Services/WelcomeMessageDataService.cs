using System;
using Nexus.Common.Interfaces.Repository;
using Nexus.Common.Messages;
using Nexus.Common.Interfaces.Services;

namespace Nexus.Service.Services {
	public class WelcomeMessageDataService : IWelcomeMessageDataService {
		private readonly IWelcomeMessageDataRepository _welcomeMessageDataRepository;

		public WelcomeMessageDataService(IWelcomeMessageDataRepository welcomeMessageDataRepository) {
			_welcomeMessageDataRepository = welcomeMessageDataRepository;
		}

		public WelcomeDataMessage GetWelcomeMessageData(Guid? userID, string language) {
			return _welcomeMessageDataRepository.GetWelcomeMessageData(userID, language);
		}
	}
}
