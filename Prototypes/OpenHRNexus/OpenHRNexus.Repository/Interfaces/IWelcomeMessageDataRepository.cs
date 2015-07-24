using System;

namespace OpenHRNexus.Repository.Interfaces {
	public interface IWelcomeMessageDataRepository {
		Guid UserId { set; }
		string Language { set; }
		string WelcomeMessageData { get; }
		DateTime LastLoginDateTime { get; }
		string SecurityGroup { get; }
	}
}
