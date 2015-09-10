using Ninject;
using Nexus.Common.Interfaces.Repository;
using Nexus.Sql_Repository;
using Nexus.Common.Interfaces;

namespace Nexus.Service.Configuration {
	public class NinjectConfig {
		public static void Config(IKernel kernel) {
			kernel.Bind<IAuthenticateRepository>().To<SqlAuthenticateRepository>();
			kernel.Bind<IWelcomeMessageDataRepository>().To<SqlAuthenticateRepository>();
			kernel.Bind<IDataRepository>().To<SqlDataRepository>();
			kernel.Bind<IEntityRepository>().To<SqlDataRepository>();
            kernel.Bind<IDictionary>().To<SqlDictionaryRepository>();
        }
    }
}
