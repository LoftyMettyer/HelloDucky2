using System;

namespace Fusion.Republisher.Core.Configuration
{
    public interface IFusionConfiguration
    {
        string Community { get; }
        bool StoreState { get; }
    }
}
