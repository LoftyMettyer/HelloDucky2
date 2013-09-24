using NServiceBus;
using NServiceBus.Faults;
using NServiceBus.Hosting.Profiles;

namespace Fusion.Connector.OpenHR
{
	class EndpointConfig : IConfigureThisEndpoint, AsA_Publisher, IWantCustomInitialization
	{
		public void Init()
		{


			Configure.With()
			         .DisableTimeoutManager()
			         .StructureMapBuilder()
			         .JsonSerializer()
			         .UnicastBus()


			         .DoNotAutoSubscribe()
                .DisableRavenInstall()

//			         .RavenPersistence("Data Source=.;Initial Catalog=OpenHR51_std;Integrated Security=True;APP=OpenHR Fusion Connector")
								//.RavenSagaPersister()
//								.RavenSubscriptionStorage()
//								.UseRavenTimeoutPersister()

				//			.DisableSecondLevelRetries()


				//		.MsmqTransport()
					.MsmqSubscriptionStorage();
				
		}
	}


	//public class MyProductionProfileHandler : IHandleProfile<Core.Production>
	//{
	//	void IHandleProfile.ProfileActivated()
	//	{
	//		if (!Configure.Instance.Configurer.HasComponent<IManageMessageFailures>())
	//			Configure.Instance.MessageForwardingInCaseOfFault();
	//	}
	//}

}




