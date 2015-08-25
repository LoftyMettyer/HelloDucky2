using System;
using Nexus.Repository.Messages;

namespace Nexus.Service.Interfaces {
	public interface IWelcomeMessageDataService {
		WelcomeDataMessage GetWelcomeMessageData(Guid? userID, string language);
	}
}