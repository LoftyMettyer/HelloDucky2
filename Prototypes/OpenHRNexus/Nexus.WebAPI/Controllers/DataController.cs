using System;
using System.Collections.Generic;
using System.Web.Http;
using Nexus.Common.Models;
using System.Security.Claims;
using Microsoft.AspNet.Identity;
using Nexus.Common.Classes;
using Nexus.Common.Classes.DataFilters;
using System.Web;
using Nexus.Common.Interfaces.Services;

namespace Nexus.WebAPI.Controllers {
//	[Authorize(Roles = "OpenHRUser")]
	public class DataController : ApiController {
		private readonly IDataService _dataService;
        //private readonly IWorkflowService _workflowService;

        private ClaimsIdentity _identity;
        private string _language;

        public DataController() {
		}

		public DataController(IDataService dataService)
		{
			_dataService = dataService;
            _identity = User.Identity as ClaimsIdentity;
            _language = HttpContext.Current.Request.UserLanguages[0].ToLowerInvariant().Trim();
        }

        public DataController(IDataService dataService, ClaimsIdentity claims, string language)
        {
            _identity = claims;
            _dataService = dataService;
            _language = language;
        }

        /// <summary>
        /// Instatiate a Process (DO WE NEED A GLOSSARY SOMEWHERE SO THIRD PARTY USERS KNOW WHAT A "PROCESS" IS?
        /// </summary>
        /// <param name="processId">Value of the process</param>
        /// <returns></returns>
        [HttpGet]
        [Authorize(Roles = "OpenHRUser")]
        public IEnumerable<WebFormModel> InstantiateProcess(int processId)
        {

            var openHRDbGuid = new Guid(_identity.GetUserId());
            List<WebFormModel> form = new List<WebFormModel>();
            WebFormModel webForm;

            if (openHRDbGuid == null || openHRDbGuid == Guid.Empty) {
                // Berties error handler goes here ?
            }
            else
            {
                webForm = _dataService.InstantiateProcess(processId, openHRDbGuid, _language);
                form.Add(webForm);
            }

            IEnumerable<WebFormModel> webFormModels = form;
            return webFormModels;

        }

        [HttpPost]
        [Authorize(Roles = "OpenHRUser")]
        public ProcessStepResponse SubmitStep(WebFormModel form)
        {
            //Guid stepId, List< KeyValuePair < int, string>> data


            // Put some clever code in an attribute extension to validate that there is a identity getuserguid?
            // Maybe this is already covered by the authorize roles = OpenHRUser?
            var userId = new Guid(_identity.GetUserId());

            return _dataService.SubmitStepForUser(form.stepid, userId, form);

        }

        [HttpGet]
        [Authorize(Roles = "OpenHRUser")]
        public IEnumerable<CalendarEventModel> GetCalendarData(string calendarType, DateTime from, DateTime to)
        {

            var userId = new Guid(_identity.GetUserId());

            var filters = new List<CalendarFilter>();
            filters.Add(new CalendarFilter() { StartRange = from, EndRange = to });

            return _dataService.GetReportData(1, filters);

        }



    }
}
