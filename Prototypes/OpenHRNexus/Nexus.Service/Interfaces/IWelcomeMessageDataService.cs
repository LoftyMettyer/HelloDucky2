﻿using System;
using Nexus.Common.Messages;

namespace Nexus.Service.Interfaces {
	public interface IWelcomeMessageDataService {
		WelcomeDataMessage GetWelcomeMessageData(Guid? userID, string language);
	}
}