using OpenHRNexus.Repository.Globals;

namespace OpenHRNexus.Repository.Configuration {
	public class EnterpriseServicesConfiguration {
		public static void Configure() {
			Global.NexusExceptionManager = new OpenHRNexus.EnterpriseServices.ExceptionHandling.NexusExceptionManager("OpenHRNexus.Repository", true, true);
			Global.NexusLoggingManager = new OpenHRNexus.EnterpriseServices.Logging.NexusLoggingManager("OpenHRNexus.Repository");
		}
	}
}
