using System;
using System.Collections.Generic;
using System.Web.Mvc;
using Ninject;
using OpenHRNexus.Service.Interfaces;
using OpenHRNexus.Service.Services;

namespace OpenHRNexus.WebAPI {
	public class NinjectConfig : IDependencyResolver {
		private readonly IKernel _kernel;

		public NinjectConfig() {
			_kernel = new StandardKernel();
			//Add bindings
			_kernel.Bind<IPersonnelRecordsService>().To<PersonnelRecordsService>();
			_kernel.Bind<ITbuserLanguagesService>().To<TbuserLanguagesService>();
			_kernel.Bind<IAuthenticateService>().To<AuthenticateService>();

			OpenHRNexus.Service.Configuration.NinjectConfig.Config(_kernel);
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
