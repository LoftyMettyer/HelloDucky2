using System;
using Nexus.Common.Messages;

namespace Nexus.Common.Interfaces.Repository {
	public interface IWelcomeMessageDataRepository {
		WelcomeDataMessage GetWelcomeMessageData(Guid? userID, string language);
	}
}
