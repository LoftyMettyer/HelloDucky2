using Ninject;
using OpenHRNexus.Repository.Interfaces;
using OpenHRNexus.Repository.SQLServer;

namespace OpenHRNexus.Service.Configuration {
	public class NinjectConfig {
		public static void Config(IKernel kernel) {
			kernel.Bind<IAuthenticateRepository>().To<SqlAuthenticateRepository>();
			kernel.Bind<IWelcomeMessageDataRepository>().To<SqlAuthenticateRepository>();
			kernel.Bind<IDataRepository>().To<SqlDataRepository>();
			kernel.Bind<IEntityRepository>().To<SqlDataRepository>();
		}
	}
}
