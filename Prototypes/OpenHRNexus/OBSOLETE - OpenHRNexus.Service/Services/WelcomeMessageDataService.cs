using System;
using OpenHRNexus.Repository.Interfaces;
using OpenHRNexus.Repository.Messages;
using OpenHRNexus.Service.Interfaces;

namespace OpenHRNexus.Service.Services {
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
