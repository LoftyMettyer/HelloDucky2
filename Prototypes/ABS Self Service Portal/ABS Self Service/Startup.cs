using Microsoft.Owin;
using Owin;

[assembly: OwinStartupAttribute(typeof(ABS_Self_Service.Startup))]
namespace ABS_Self_Service
{
    public partial class Startup
    {
        public void Configuration(IAppBuilder app)
        {
            ConfigureAuth(app);
        }
    }
}
