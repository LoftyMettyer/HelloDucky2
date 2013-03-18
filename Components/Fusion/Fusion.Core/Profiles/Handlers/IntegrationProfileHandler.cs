using System;
using System.Collections.Generic;
using System.Linq;

namespace Fusion.Core.Profiles.Handlers
{
    using NServiceBus.Saga;
    using NServiceBus.Unicast.Subscriptions;
    using NServiceBus.Hosting.Profiles;
    using NServiceBus;
    using NServiceBus.Faults;
    using NServiceBus.Hosting.Windows.Profiles.Handlers;

    internal class IntegrationProfileHandler : IHandleProfile<Fusion.Core.Integration>, IWantTheEndpointConfig, IWantTheListOfActiveProfiles
    {
        void IHandleProfile.ProfileActivated()
        {
            Configure.Instance.RavenPersistence();

            if (!Configure.Instance.Configurer.HasComponent<IManageMessageFailures>())
                Configure.Instance.MessageForwardingInCaseOfFault();

            if (!Configure.Instance.Configurer.HasComponent<ISagaPersister>())
                Configure.Instance.RavenSagaPersister();


            if (Config is AsA_Publisher && !Configure.Instance.Configurer.HasComponent<ISubscriptionStorage>())
            {
                if ((ActiveProfiles.Contains(typeof(Master))) || (ActiveProfiles.Contains(typeof(Worker))) || (ActiveProfiles.Contains(typeof(Distributor))))
                    Configure.Instance.RavenSubscriptionStorage();
                else
                    Configure.Instance.MsmqSubscriptionStorage();
            }

            WindowsInstallerRunner.RunInstallers = true;
            WindowsInstallerRunner.RunInfrastructureInstallers = false;
        }

        public IConfigureThisEndpoint Config { get; set; }

        public IEnumerable<Type> ActiveProfiles { get; set; }
    }
}