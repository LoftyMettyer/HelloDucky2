using System;
using OpenHRNexus.Repository.Messages;

namespace OpenHRNexus.Repository.Interfaces {
	public interface IWelcomeMessageDataRepository {
		WelcomeDataMessage GetWelcomeMessageData(Guid? userID, string language);
	}
}
