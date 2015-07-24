using System;

namespace OpenHRNexus.Repository.Interfaces
{
	public interface IWelcomeMessageData
	{
		Guid UserID { set; }
		string Language { set; }
		string WelcomeMessageData { get; }
		DateTime LastLoginDateTime { get; }
		string SecurityGroup { get; }
	}
}
