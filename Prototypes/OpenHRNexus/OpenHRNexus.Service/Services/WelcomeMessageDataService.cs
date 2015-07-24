using System;
using OpenHRNexus.Repository.Interfaces;
using OpenHRNexus.Service.Interfaces;

namespace OpenHRNexus.Service.Services {
	public class WelcomeMessageDataService : IWelcomeMessageDataService {
		private readonly IWelcomeMessageDataRepository _welcomeMessageDataRepository;

		public WelcomeMessageDataService(IWelcomeMessageDataRepository welcomeMessageDataRepository) {
			_welcomeMessageDataRepository = welcomeMessageDataRepository;
		}

		public string WelcomeMessageData {
			get { return _welcomeMessageDataRepository.WelcomeMessageData; }
		}
		public DateTime LastLoginDateTime {
			get { return _welcomeMessageDataRepository.LastLoginDateTime; }
		}
		public string SecurityGroup {
			get { return _welcomeMessageDataRepository.SecurityGroup; }
		}
	}
}
