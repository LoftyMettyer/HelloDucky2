using OpenHRNexus.Service.Globals;

namespace OpenHRNexus.Service.Configuration {
	public class EnterpriseServicesConfiguration {
		public static void Configure() {
			Global.NexusExceptionManager = new OpenHRNexus.EnterpriseServices.ExceptionHandling.NexusExceptionManager("OpenHRNexus.Service", true, true);
			Global.NexusLoggingManager = new OpenHRNexus.EnterpriseServices.Logging.NexusLoggingManager("OpenHRNexus.Service");
		}
	}
}
