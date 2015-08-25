using System;
using System.Collections.Generic;
using System.Web.Mvc;
using Ninject;
using Nexus.Repository.Interfaces;
using Nexus.Service.Interfaces;
using Nexus.Service.Services;

namespace Nexus.WebAPI {
	public class NinjectConfig : IDependencyResolver {
		private readonly IKernel _kernel;

		public NinjectConfig() {
			_kernel = new StandardKernel();
			//Add bindings
			_kernel.Bind<IAuthenticateService>().To<AuthenticateService>();
			_kernel.Bind<IWelcomeMessageDataService>().To<WelcomeMessageDataService>();
			_kernel.Bind<IDataService>().To<DataService>();
			_kernel.Bind<IEntityService>().To<EntityService>();


			Nexus.Service.Configuration.NinjectConfig.Config(_kernel);
		}

		public IKernel Kernel { get { return _kernel; } }

		public object GetService(Type serviceType) {
			return _kernel.TryGet(serviceType);
		}

		public IEnumerable<object> GetServices(Type serviceType) {
			return _kernel.GetAll(serviceType);
		}
	}
}
