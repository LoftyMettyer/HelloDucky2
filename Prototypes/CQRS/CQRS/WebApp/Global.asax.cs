using System.Reflection;
using System.Web;
using System.Web.Http;
using System.Web.Mvc;
using System.Web.Optimization;
using System.Web.Routing;
using Autofac;
using Infra.IOCContainer;
using Infra.Processor;

namespace WebApp
{
	// Note: For instructions on enabling IIS6 or IIS7 classic mode, 
	// visit http://go.microsoft.com/?LinkId=9394801

	public class MvcApplication : HttpApplication
	{
		private ContainerBuilder builder;
		private IContainer container;

		protected void Application_Start()
		{
			AreaRegistration.RegisterAllAreas();

			WebApiConfig.Register(GlobalConfiguration.Configuration);
			FilterConfig.RegisterGlobalFilters(GlobalFilters.Filters);
			RouteConfig.RegisterRoutes(RouteTable.Routes);
			BundleConfig.RegisterBundles(BundleTable.Bundles);
			AuthConfig.RegisterAuth();
			DependecyResolver();
		}

		private void DependecyResolver()
		{
			builder = IoCBuilder.ContainerBuilder;

			//controllers
			//builder.RegisterControllers(Assembly.GetExecutingAssembly());
			//builder.RegisterApiControllers(Assembly.GetExecutingAssembly());

			Assembly assembly = typeof (QueryProcessor).Assembly;
			//builder.Register<IDbContext>(c => new ApplicationContext()).InstancePerLifetimeScope();
			builder.RegisterAssemblyTypes(assembly).AsImplementedInterfaces().InstancePerLifetimeScope();
			RegisterGenericInterfaces(builder, assembly);


			container = IoCBuilder.Container;

			//DependencyResolver.SetResolver(new AutofacDependencyResolver(container));
			//GlobalConfiguration.Configuration.DependencyResolver = new AutofacWebApiDependencyResolver(container);
		}

		private void RegisterGenericInterfaces(ContainerBuilder builder, Assembly assembly)
		{
			//foreach (var type in assembly.GetTypes().Where(m => m.GetInterfaces().Count() > 0))
			//{
			//    var repositoryInterfaceType = type.GetInterfaces().Where(m => m.Name == typeof(IRepository<>).Name);

			//    if (repositoryInterfaceType != null && repositoryInterfaceType.Count() > 0)
			//    {
			//        builder.RegisterGeneric(type).As(typeof(IRepository<>)).InstancePerLifetimeScope();
			//    }
			//}
		}
	}
}