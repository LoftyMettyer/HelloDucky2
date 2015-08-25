using Ninject;
using Nexus.Repository.Interfaces;
using Nexus.Repository.SQLServer;

namespace Nexus.Service.Configuration {
	public class NinjectConfig {
		public static void Config(IKernel kernel) {
			kernel.Bind<IAuthenticateRepository>().To<SqlAuthenticateRepository>();
			kernel.Bind<IWelcomeMessageDataRepository>().To<SqlAuthenticateRepository>();
			kernel.Bind<IDataRepository>().To<SqlDataRepository>();
			kernel.Bind<IEntityRepository>().To<SqlDataRepository>();
		}
	}
}
