using Nexus.Sql_Repository.Globals;

namespace Nexus.Sql_Repository.Configuration {
	public class EnterpriseServicesConfiguration {
		public static void Configure() {
			Global.NexusExceptionManager = new Nexus.EnterpriseService.ExceptionHandling.NexusExceptionManager("Nexus.SqlRepository", true, true);
			Global.NexusLoggingManager = new Nexus.EnterpriseService.Logging.NexusLoggingManager("Nexus.SqlRepository");
		}
	}
}
