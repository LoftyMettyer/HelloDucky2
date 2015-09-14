//using System;
//using System.Collections.Generic;
//using System.Web.Http.Dependencies;
//using Ninject;
//using Nexus.Repository.Interfaces;
//using Nexus.Repository.SQLServer;
//using Nexus.Service.Interfaces;
//using Nexus.Service.Services;

//namespace Nexus.WebAPI.Tests.Configuration {

//	public class NinjectConfig : IDependencyResolver
//	{
//		private readonly IKernel _kernel;

//		public NinjectConfig()
//		{
//			_kernel = new StandardKernel();
//			_kernel.Bind<IDataService>().To<DataService>();
//			_kernel.Bind<IProcessRepository>().To<SqlProcessRepository>();

//			Nexus.Service.Configuration.NinjectConfig.Config(_kernel);
//		}

//		public IKernel Kernel { get { return _kernel; } }

//		public object GetService(Type serviceType)
//		{
//			return _kernel.TryGet(serviceType);
//		}

//		public IEnumerable<object> GetServices(Type serviceType)
//		{
//			return _kernel.GetAll(serviceType);
//		}

//		public IDependencyScope BeginScope()
//		{
//			throw new NotImplementedException();
//		}

//		public void Dispose()
//		{
//			throw new NotImplementedException();
//		}
//	}

//}
