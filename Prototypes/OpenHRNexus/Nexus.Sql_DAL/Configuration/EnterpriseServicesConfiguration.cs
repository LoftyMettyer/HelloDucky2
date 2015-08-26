using Nexus.Repository.Globals;

namespace Nexus.Repository.Configuration {
	public class EnterpriseServicesConfiguration {
		public static void Configure() {
			Global.NexusExceptionManager = new Nexus.EnterpriseServices.ExceptionHandling.NexusExceptionManager("Nexus.Repository", true, true);
			Global.NexusLoggingManager = new Nexus.EnterpriseServices.Logging.NexusLoggingManager("Nexus.Repository");
		}
	}
}
