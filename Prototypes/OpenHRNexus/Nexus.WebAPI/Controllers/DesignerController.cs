using Nexus.Common.Enums;
using Nexus.Common.Interfaces.Services;
using System.Security.Claims;
using System.Web;
using System.Web.Http;

namespace Nexus.WebAPI.Controllers
{
    /// <summary>
    /// This controller is used for accessing the nexus metadata design services
    /// </summary>
    public class DesignerController : ApiController
    {

        private readonly IDesignerService _designerService;
        private ClaimsIdentity _identity;
        private string _language;

        /// <summary>
        /// Loads the design contoller
        /// </summary>
        /// <param name="designerService"></param>
        public DesignerController(IDesignerService designerService)
        {
            _designerService = designerService;
            _identity = User.Identity as ClaimsIdentity;
            _language = HttpContext.Current.Request.UserLanguages[0].ToLowerInvariant().Trim();
        }

        /// <summary>
        /// Loads the design controller with parameters from the unit test projects
        /// </summary>
        /// <param name="designerService"></param>
        /// <param name="claims"></param>
        /// <param name="language"></param>
        public DesignerController(IDesignerService designerService, ClaimsIdentity claims, string language)
        {
            _designerService = designerService;
            _identity = claims;
            _language = language;
        }

        public void AddEntity(EntityType type, string name)
        {

        }

    }
}