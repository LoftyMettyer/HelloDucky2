using NServiceBus;
using Subscriber1;

namespace SimpleConnector
{
    class EndpointConfig : IConfigureThisEndpoint, AsA_Server, IWantCustomInitialization {

        public void Init()
        {
            NServiceBus.Configure.With()
                .StructureMapBuilder()
                .JsonSerializer()
                .UnicastBus();

            NServiceBus.Configure.Instance.Configurer.ConfigureComponent<OnHoldMessageMutator>(DependencyLifecycle.InstancePerCall);
        }
    }
}
