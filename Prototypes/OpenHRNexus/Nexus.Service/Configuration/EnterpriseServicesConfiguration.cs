using Nexus.Service.Globals;

namespace Nexus.Service.Configuration {
	public class EnterpriseServicesConfiguration {
		public static void Configure() {
			Global.NexusExceptionManager = new Nexus.EnterpriseService.ExceptionHandling.NexusExceptionManager("Nexus.Service", true, true);
			Global.NexusLoggingManager = new Nexus.EnterpriseService.Logging.NexusLoggingManager("Nexus.Service");
		}
	}
}
