using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Security.Claims;
using System.Web;
using System.Web.Http;
using OpenHRNexus.Service.Interfaces;
using OpenHRNexus.WebAPI.Resources;
using Microsoft.AspNet.Identity;

namespace OpenHRNexus.WebAPI.Controllers
{
    public class ResourceController : ApiController
    {
        private readonly IWelcomeMessageDataService _welcomeMessageDataService;

        public ResourceController()
        {
        }

        public ResourceController(IWelcomeMessageDataService welcomeMessageDataService)
        {
            _welcomeMessageDataService = welcomeMessageDataService;
        }

        [HttpGet]
        public IEnumerable<KeyValuePair<string, string>> GetResourceValues([FromUri] List<string> parameter)
        {
            return parameter.ToDictionary(s => s, s => Resource.ResourceManager.GetString(s));
        }

        [HttpGet]
        [Authorize(Roles = "OpenHRUser")]
        public IEnumerable<string> GetProtectedResourceValue(string resource)
        {
            // TODO - Investigate whether this is the best way to interrogate languages - performance hit?
            var language = "EN-GB";
            if (HttpContext.Current.Request.UserLanguages != null)
            {
                language = HttpContext.Current.Request.UserLanguages[0].ToLowerInvariant().Trim();
            }

            //Get the OpenHR guid out of the jwt
            var identity = User.Identity as ClaimsIdentity;

            if (identity != null)
            {
                string openHRDbGuid = User.Identity.GetUserId();

                var welcomeMessage = _welcomeMessageDataService.GetWelcomeMessageData(new Guid(openHRDbGuid), language);

                var translation = Resource.ResourceManager.GetString(resource);
                if (translation != null)
                    return new[]
                    {
                            translation
                                .Replace("#FullName#", welcomeMessage.Message)
                                .Replace("#LastLoginDate#", welcomeMessage.LastLoggedOn.ToString(CultureInfo.CurrentCulture))
                                .Replace("#SecurityGroup#", welcomeMessage.SecurityGroup)
                        };
            }

            return new[] { "Welcome." };
        }

        [HttpGet]
        [Authorize(Roles = "OpenHRUser")]
        public IEnumerable<string> InstantiateProcess(string InstanceId, string ElementId, bool NewRecord)
        {
            return new[] { "1", "2" };
        }

    }
}
