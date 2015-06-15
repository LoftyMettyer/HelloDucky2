using System.Collections.Generic;
using System.Web.Http;
using OpenHRNexus.Common.Models;
using OpenHRNexus.Service.Interfaces;

namespace OpenHRNexus.WebAPI.Controllers {
	public class PersonnelRecordsController : ApiController {
		private readonly IPersonnelRecordsService _personnelRecordsService;

		public PersonnelRecordsController(IPersonnelRecordsService personnelRecordsService) {
			_personnelRecordsService = personnelRecordsService;
		}

		public PersonnelRecordsController() {
		}

		public IEnumerable<Personnel_Records_Model> Get() {
			return _personnelRecordsService.List();
		}
	}
}