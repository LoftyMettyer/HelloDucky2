using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Web.Http;
using OpenHRNexus.Common.Models;
using OpenHRNexus.Service.Interfaces;
using OpenHRNexus.WebAPI.Extensions;

namespace OpenHRNexus.WebAPI.Controllers {
//	[Authorize(Roles = "OpenHRUser")]
	public class DataController : ApiController {
		private readonly IDataService _dataService;
        //private readonly IWorkflowService _workflowService;

		public DataController() {
		}

		public DataController(IDataService dataService)
		{
			_dataService = dataService;
		}

		[HttpGet]
		public string GetReportData(string id)
		{
			int dataId;

			var result = Int32.TryParse(id, out dataId);

			if (result)
			{
				var data = _dataService.GetData(dataId);
				return data.ToJsonResult().ToString();
			}
			else
			{
				var data = _dataService.GetData();
				return data.ToJsonResult().ToString();
			}

		}

        [HttpGet]
        [Authorize(Roles = "OpenHRUser")]
        public IEnumerable<WebFormModel> InstantiateProcess(int instanceId, int elementId, bool newRecord)
        {

            var fields = _dataService.GetWebFormFields(elementId);

            //            fields.Translate("en-gb");

            List<WebFormModel> form = new List<WebFormModel>();
            form.Add(new WebFormModel
            {
                form_id = "1",
                form_name = "Test Form",
                form_fields = fields.ToList()
            });

            IEnumerable<WebFormModel> webFormModels = form;

            return webFormModels;

        }



    }
}
