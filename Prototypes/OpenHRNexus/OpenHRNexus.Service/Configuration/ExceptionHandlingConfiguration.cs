using OpenHRNexus.Service.Globals;

namespace OpenHRNexus.Service.Configuration {
	public class ExceptionHandlingConfiguration {
		public static void Configure() {
			Global.NexusExceptionManager = new OpenHRNexus.EnterpriseServices.ExceptionHandling.NexusExceptionManager("OpenHRNexus.Service", true, true);
		}
	}
}
