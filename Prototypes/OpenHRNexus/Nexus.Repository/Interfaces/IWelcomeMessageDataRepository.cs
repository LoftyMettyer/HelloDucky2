using System;
using Nexus.Repository.Messages;

namespace Nexus.Repository.Interfaces {
	public interface IWelcomeMessageDataRepository {
		WelcomeDataMessage GetWelcomeMessageData(Guid? userID, string language);
	}
}
