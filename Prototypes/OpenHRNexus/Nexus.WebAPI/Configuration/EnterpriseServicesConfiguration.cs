using Nexus.WebAPI.Globals;

namespace Nexus.WebAPI.Configuration {
	public class EnterpriseServicesConfiguration {
		public static void Configure() {
			Global.NexusExceptionManager = new Nexus.EnterpriseServices.ExceptionHandling.NexusExceptionManager("Nexus.WebAPI", true, true);
			Global.NexusLoggingManager = new Nexus.EnterpriseServices.Logging.NexusLoggingManager("Nexus.WebAPI");
		}
	}
}
