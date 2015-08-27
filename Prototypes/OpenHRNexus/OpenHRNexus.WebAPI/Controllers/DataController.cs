using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Web.Http;
using OpenHRNexus.Common.Models;
using OpenHRNexus.Service.Interfaces;
using OpenHRNexus.WebAPI.Extensions;
using System.Security.Claims;
using Microsoft.AspNet.Identity;

namespace OpenHRNexus.WebAPI.Controllers {
//	[Authorize(Roles = "OpenHRUser")]
	public class DataController : ApiController {
		private readonly IBusinessProcessService _dataService;
        //private readonly IWorkflowService _workflowService;

		public DataController() {
		}

		public DataController(IBusinessProcessService dataService)
		{
			_dataService = dataService;
		}

		[HttpGet]
		public string GetReportData(string id)
		{
            //int dataId;

            //var result = Int32.TryParse(id, out dataId);

            //if (result)
            //{
            //	var data = _dataService.GetData(dataId);
            //	return data.ToJsonResult().ToString();
            //}
            //else
            //{
            //	var data = _dataService.GetData();
            //	return data.ToJsonResult().ToString();
            //}
            return "notyetimplemented";

		}

        [HttpGet]
        [Authorize(Roles = "OpenHRUser")]
        public IEnumerable<WebFormModel> InstantiateProcess(int instanceId, int processId, bool newRecord)
        {

            // TODO - This bit needs to extract from the JWT
            //var identity = User.Identity as ClaimsIdentity;
            //var openHRDbGuid = new Guid(identity.GetUserId());
            var openHRDbGuid = new Guid("088C6A78-E14A-41B0-AD93-4FB7D3ADE96C");

            //var webForm = _dataService.GetWebForm(elementId, openHRDbGuid);

            var webForm = _dataService.GetWebFormForProcessAndUser(processId, openHRDbGuid);


            List<WebFormModel> form = new List<WebFormModel>();
            form.Add(webForm);

            IEnumerable<WebFormModel> webFormModels = form;
            return webFormModels;

        }

    }
}
