using System;
using System.Collections.Generic;
using System.Linq;
using System.Web.Http;
using OpenHRNexus.Common.Models;
using OpenHRNexus.Service.Interfaces;

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

		[HttpGet]
		public IEnumerable<DynamicDataModel> GetReportData(string id)
		{
			int dataId;

			var result = Int32.TryParse(id, out dataId);

			if (result)
			{
				return _dataService.GetData(dataId);
			}
			else
			{
				return _dataService.GetData();
			}

		}
	}
}
