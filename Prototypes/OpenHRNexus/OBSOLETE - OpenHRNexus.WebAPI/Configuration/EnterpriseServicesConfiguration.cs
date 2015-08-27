using OpenHRNexus.WebAPI.Globals;

namespace OpenHRNexus.WebAPI.Configuration {
	public class EnterpriseServicesConfiguration {
		public static void Configure() {
			Global.NexusExceptionManager = new OpenHRNexus.EnterpriseServices.ExceptionHandling.NexusExceptionManager("OpenHRNexus.WebAPI", true, true);
			Global.NexusLoggingManager = new OpenHRNexus.EnterpriseServices.Logging.NexusLoggingManager("OpenHRNexus.WebAPI");
		}
	}
}
