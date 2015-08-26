using System.Web.Http;
using System.Web.Mvc;
using System.Web.Optimization;
using System.Web.Routing;
using Nexus.WebAPI.Handlers;

namespace Nexus.WebAPI {
	public class WebApiApplication : System.Web.HttpApplication {
		protected void Application_Start() {
			AreaRegistration.RegisterAllAreas();
			GlobalConfiguration.Configure(WebApiConfig.Register);
			FilterConfig.RegisterGlobalFilters(GlobalFilters.Filters);
			RouteConfig.RegisterRoutes(RouteTable.Routes);
			BundleConfig.RegisterBundles(BundleTable.Bundles);

			Service.Configuration.AutoMapperConfig.Configure();

			var ninjectConfig = new NinjectConfig();
			DependencyResolver.SetResolver(ninjectConfig);
			GlobalConfiguration.Configuration.DependencyResolver = new NinjectResolver(ninjectConfig.Kernel);

			//Add the message logging handler to the message handlers collection
			GlobalConfiguration.Configuration.MessageHandlers.Add(new MessageLoggingHandler());

			//Initialise Exception Handling for all projects in solution
			Nexus.WebAPI.Configuration.EnterpriseServicesConfiguration.Configure();
			Nexus.Service.Configuration.EnterpriseServicesConfiguration.Configure();
			Nexus.Sql_Repository.Configuration.EnterpriseServicesConfiguration.Configure();
		}
	}
}
