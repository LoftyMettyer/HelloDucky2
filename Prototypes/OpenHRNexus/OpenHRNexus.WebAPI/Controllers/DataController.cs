using System;
using System.Collections.Generic;
using System.Linq;
using System.Web.Http;
using OpenHRNexus.Common.Models;
using OpenHRNexus.Service.Interfaces;

namespace OpenHRNexus.WebAPI.Controllers {
//	[Authorize(Roles = "ApplicationUser")]
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
			return _dataService.GetData(Convert.ToInt32(id));
		}
	}
}
