using Ninject;
using OpenHRNexus.Repository.Interfaces;
using OpenHRNexus.Repository.Repositories.SQLServer;
using OpenHRNexus.Repository.Repositories.MySQL;

namespace OpenHRNexus.Service.Configuration {
	public class NinjectConfig {
		public static void Config(IKernel kernel) {
			kernel.Bind<IPersonnelRecordsRepository>().To<SQLPersonnelRecordsRepository>();
			kernel.Bind<Itbuser_LanguagesRepository>().To<MySQLtbuser_LanguagesRepository>();
		}
	}
}
