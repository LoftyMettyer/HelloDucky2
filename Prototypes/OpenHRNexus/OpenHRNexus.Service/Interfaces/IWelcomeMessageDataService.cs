using System;

namespace OpenHRNexus.Service.Interfaces {
	public interface IWelcomeMessageDataService {
		string WelcomeMessageData { get; }
		DateTime LastLoginDateTime { get; }
		string SecurityGroup { get; }
	}
}