using System;
using Nexus.Repository.Interfaces;
using Nexus.Repository.Messages;
using Nexus.Service.Interfaces;

namespace Nexus.Service.Services {
	public class WelcomeMessageDataService : IWelcomeMessageDataService {
		private readonly IWelcomeMessageDataRepository _welcomeMessageDataRepository;

		public WelcomeMessageDataService(IWelcomeMessageDataRepository welcomeMessageDataRepository) {
			_welcomeMessageDataRepository = welcomeMessageDataRepository;
		}

		public WelcomeDataMessage GetWelcomeMessageData(Guid? userID, string language)
		{
			return _welcomeMessageDataRepository.GetWelcomeMessageData(userID, language);
		}
	}
}
