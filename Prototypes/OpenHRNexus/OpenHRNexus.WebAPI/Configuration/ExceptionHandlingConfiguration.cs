using OpenHRNexus.WebAPI.Globals;

namespace OpenHRNexus.WebAPI.Configuration {
	public class ExceptionHandlingConfiguration {
		public static void Configure() {
			Global.NexusExceptionManager = new OpenHRNexus.EnterpriseServices.ExceptionHandling.NexusExceptionManager("OpenHRNexus.WebAPI", true, true);
		}
	}
}
