using System;
using OpenHRNexus.Repository.Messages;

namespace OpenHRNexus.Service.Interfaces {
	public interface IWelcomeMessageDataService {
		WelcomeDataMessage GetWelcomeMessageData(Guid? userID, string language);
	}
}