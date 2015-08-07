using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Web.Http;
using System.Web.Mvc;
using OpenHRNexus.Common.Models;
using OpenHRNexus.Service.Interfaces;
using OpenHRNexus.WebAPI.Extensions;

namespace OpenHRNexus.WebAPI.Controllers {
//	[Authorize(Roles = "OpenHRUser")]
	public class DataController : ApiController {
		private readonly IDataService _dataService;

		public DataController() {
		}

		public DataController(IDataService dataService)
		{
			_dataService = dataService;
		}

		[System.Web.Http.HttpGet]
		public MvcHtmlString GetReportData(string id)
		{
			int dataId;

			var result = Int32.TryParse(id, out dataId);

			if (result)
			{
				var data = _dataService.GetData(dataId);
				return data.ToJsonResult();
			}
			else
			{
				var data = _dataService.GetData();
				return data.ToJsonResult();
			}

		}
	}
}
