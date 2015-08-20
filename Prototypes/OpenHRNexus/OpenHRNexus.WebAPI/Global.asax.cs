using System.Web.Http;
using System.Web.Mvc;
using System.Web.Optimization;
using System.Web.Routing;

namespace OpenHRNexus.WebAPI {
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

			//Initialise Exception Handling for all projects in solution
			OpenHRNexus.WebAPI.Configuration.ExceptionHandlingConfiguration.Configure();
			OpenHRNexus.Service.Configuration.ExceptionHandlingConfiguration.Configure();
			OpenHRNexus.Repository.Configuration.ExceptionHandlingConfiguration.Configure();
		}
	}
}
