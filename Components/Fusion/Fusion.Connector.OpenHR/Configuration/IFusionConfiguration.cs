﻿namespace Fusion.Connector.OpenHR.Configuration
{
	public interface IFusionConfiguration
	{
		string ServiceName { get; }
		string Community { get; }
		string SendAsUser { get; }
	}
}
