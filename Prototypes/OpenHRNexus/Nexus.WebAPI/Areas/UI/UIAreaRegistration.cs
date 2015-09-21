using System.Web.Mvc;

namespace Nexus.WebAPI.Areas.UI
{
    public class UIAreaRegistration : AreaRegistration 
    {
        public override string AreaName 
        {
            get 
            {
                return "UI";
            }
        }

        public override void RegisterArea(AreaRegistrationContext context) 
        {
            context.MapRoute(
                "UI_default",
                "UI/{controller}/{action}/{id}",
                new { action = "Index", id = UrlParameter.Optional }
            );
        }
    }
}