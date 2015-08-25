using Nexus.Service.Globals;

namespace Nexus.Service.Configuration {
	public class EnterpriseServicesConfiguration {
		public static void Configure() {
			Global.NexusExceptionManager = new Nexus.EnterpriseServices.ExceptionHandling.NexusExceptionManager("Nexus.Service", true, true);
			Global.NexusLoggingManager = new Nexus.EnterpriseServices.Logging.NexusLoggingManager("Nexus.Service");
		}
	}
}
