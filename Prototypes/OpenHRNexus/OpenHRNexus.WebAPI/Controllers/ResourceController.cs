using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Security.Claims;
using System.Web;
using System.Web.Http;
using Microsoft.AspNet.Identity;
using OpenHRNexus.Common.Models;
using OpenHRNexus.Service.Interfaces;
using OpenHRNexus.WebAPI.Resources;

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
        public IEnumerable<WebFormModel> InstantiateProcess(string instanceId, string elementId, bool newRecord)
        {

            List<WebFormModel> form = new List<WebFormModel>();

            List<WebFormFields> fields = new List<WebFormFields>
            {
                new WebFormFields
                {
                    field_id = 1,
                    field_title = "First Name",
                    field_type = "textfield",
                    field_value = "John",
                    field_required = true,
                    field_disabled = false
                },
                new WebFormFields
                {
                    field_id = 2,
                    field_title = "Last Name",
                    field_type = "textfield",
                    field_value = "Doe",
                    field_required = true,
                    field_disabled = false
                }
            };



            List<WebFormFieldOption> options = new List<WebFormFieldOption>
            {
                new WebFormFieldOption
                {
                    option_id = 1,
                    option_title = "Male",
                    option_value = 1
                },
                new WebFormFieldOption
                {
                    option_id = 2,
                    option_title = "Female",
                    option_value = 2
                }
            };

            fields.Add(new WebFormFields
            {
                field_id = 3,
                field_title = "Gender",
                field_type = "radio",
                field_value = "2",
                field_required = true,
                field_disabled = false,
                field_options = options
            });

            fields.Add(new WebFormFields
            {
                field_id = 4,
                field_title = "Email Address",
                field_type = "email",
                field_value = "test@example.com",
                field_required = true,
                field_disabled = false
            });

            fields.Add(new WebFormFields
            {
                field_id = 5,
                field_title = "Password",
                field_type = "password",
                field_value = "",
                field_required = true,
                field_disabled = false
            });

            fields.Add(new WebFormFields
            {
                field_id = 6,
                field_title = "Birth Date",
                field_type = "date",
                field_value = "17.09.1971",
                field_required = true,
                field_disabled = false
            });

            options = new List<WebFormFieldOption>
            {
                new WebFormFieldOption
                {
                    option_id = 1,
                    option_title = "--Please Select--",
                    option_value = 1
                },
                new WebFormFieldOption
                {
                    option_id = 2,
                    option_title = "Internet Explorer",
                    option_value = 2
                },
                new WebFormFieldOption
                {
                    option_id = 3,
                    option_title = "Google Chrome",
                    option_value = 3
                },
                new WebFormFieldOption
                {
                    option_id = 4,
                    option_title = "Mozilla Firefox",
                    option_value = 4
                }
            };

            fields.Add(new WebFormFields
            {
                field_id = 7,
                field_title = "Your browser",
                field_type = "dropdown",
                field_value = "2",
                field_required = false,
                field_disabled = false,
                field_options = options
            });

            fields.Add(new WebFormFields
            {
                field_id = 8,
                field_title = "Additional Comments",
                field_type = "textarea",
                field_value = "Please type here...",
                field_required = false,
                field_disabled = false
            });

            fields.Add(new WebFormFields
            {
                field_id = 9,
                field_title = "I accept the terms and conditions",
                field_type = "checkbox",
                field_value = "0",
                field_required = true,
                field_disabled = false
            });

            fields.Add(new WebFormFields
            {
                field_id = 10,
                field_title = "I have a secret",
                field_type = "hidden",
                field_value = "X",
                field_required = false,
                field_disabled = false
            });

            form.Add(new WebFormModel
            {
                form_id = "1",
                form_name = "Test Form",
                form_fields = fields
            });

            IEnumerable<WebFormModel> webFormModels = form;

            return webFormModels;

        }

    }
}
