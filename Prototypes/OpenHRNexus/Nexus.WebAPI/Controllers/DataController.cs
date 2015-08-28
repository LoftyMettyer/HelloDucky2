using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Web.Http;
using Nexus.Common.Models;
using Nexus.Service.Interfaces;
using Nexus.WebAPI.Extensions;
using System.Security.Claims;
using Microsoft.AspNet.Identity;

namespace Nexus.WebAPI.Controllers {
//	[Authorize(Roles = "OpenHRUser")]
	public class DataController : ApiController {
		private readonly IDataService _dataService;
        //private readonly IWorkflowService _workflowService;

        private ClaimsIdentity _identity;

        public DataController() {
		}

		public DataController(IDataService dataService)
		{
			_dataService = dataService;
            _identity = User.Identity as ClaimsIdentity;
		}

        public DataController(IDataService dataService, ClaimsIdentity claims)
        {
            _identity = claims;
            _dataService = dataService;
        }


        [HttpGet]
        [Authorize(Roles = "OpenHRUser")]
        public IEnumerable<WebFormModel> InstantiateProcess(int instanceId, int elementId, bool newRecord)
        {

            var openHRDbGuid = new Guid(_identity.GetUserId());
            List<WebFormModel> form = new List<WebFormModel>();
            WebFormModel webForm;

            if (openHRDbGuid == null || openHRDbGuid == Guid.Empty) {
                // Berties error handler goes here ?
            }
            else
            {
                webForm = _dataService.GetWebForm(elementId, openHRDbGuid);
                form.Add(webForm);

            }

            IEnumerable<WebFormModel> webFormModels = form;
            return webFormModels;

        }

    }
}
