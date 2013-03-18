
namespace Fusion.Core.Profiles.Handlers
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using NServiceBus;
    using NServiceBus.Faults;
    using NServiceBus.Hosting.Profiles;
    using NServiceBus.Saga;
    using NServiceBus.Unicast.Subscriptions;

    internal class ProductionProfileHandler : IHandleProfile<Fusion.Core.Production>, IWantTheEndpointConfig, IWantTheListOfActiveProfiles
    {
        void IHandleProfile.ProfileActivated()
        {
            Configure.Instance.RavenPersistence();

            if (!Configure.Instance.Configurer.HasComponent<ISagaPersister>())
                Configure.Instance.RavenSagaPersister();

            if (!Configure.Instance.Configurer.HasComponent<IManageMessageFailures>())
                Configure.Instance.MessageForwardingInCaseOfFault();

            if (Config is AsA_Publisher && !Configure.Instance.Configurer.HasComponent<ISubscriptionStorage>())
            {
                if ((ActiveProfiles.Contains(typeof(Master))) || (ActiveProfiles.Contains(typeof(Worker))) || (ActiveProfiles.Contains(typeof(Distributor))))
                    Configure.Instance.RavenSubscriptionStorage();
                else
                    Configure.Instance.MsmqSubscriptionStorage();


            }
        }

        public IConfigureThisEndpoint Config { get; set; }

        public IEnumerable<Type> ActiveProfiles { get; set; }
    }
}