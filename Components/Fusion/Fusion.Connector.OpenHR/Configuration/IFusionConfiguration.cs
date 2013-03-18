using System;

namespace Fusion.Connector.OpenHR.Configuration
{
    public interface IFusionConfiguration
    {
        string ServiceName { get; }
        string InputQueue { get; }
    }
}
