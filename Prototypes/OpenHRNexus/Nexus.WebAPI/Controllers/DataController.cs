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
			_language = "en-gb";
			if (HttpContext.Current.Request.UserLanguages != null) {
				_language = HttpContext.Current.Request.UserLanguages[0].ToLowerInvariant().Trim();
			}
		}

        public DataController(IDataService dataService, ClaimsIdentity claims, string language)
        {
            _identity = claims;
            _dataService = dataService;
            _language = language;
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
