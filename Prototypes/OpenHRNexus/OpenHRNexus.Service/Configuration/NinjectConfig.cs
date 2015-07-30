using Ninject;
using OpenHRNexus.Repository.Interfaces;
using OpenHRNexus.Repository.MockRepository;
using OpenHRNexus.Repository.MySQL;
using OpenHRNexus.Repository.SQLServer;

namespace OpenHRNexus.Service.Configuration {
	public class NinjectConfig {
		public static void Config(IKernel kernel) {
			kernel.Bind<IPersonnelRecordsRepository>().To<SqlPersonnelRecordsRepository>();
			kernel.Bind<ITbuserLanguagesRepository>().To<MySqlTbuserLanguagesRepository>();
			kernel.Bind<IAuthenticateRepository>().To<SqlAuthenticateRepository>();
			kernel.Bind<IWelcomeMessageDataRepository>().To<SqlAuthenticateRepository>();
			kernel.Bind<IDataRepository>().To<MockDataRepository>();
		}
	}
}
