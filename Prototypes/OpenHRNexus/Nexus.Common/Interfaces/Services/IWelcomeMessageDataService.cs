using System;
using Nexus.Common.Messages;

namespace Nexus.Common.Interfaces.Services {
	public interface IWelcomeMessageDataService {
		WelcomeDataMessage GetWelcomeMessageData(Guid? userID, string language);
	}
}