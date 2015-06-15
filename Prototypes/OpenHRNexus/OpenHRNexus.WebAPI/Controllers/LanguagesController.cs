using System.Collections.Generic;
using System.Web.Http;
using OpenHRNexus.Common.Models;
using OpenHRNexus.Service.Interfaces;

namespace OpenHRNexus.WebAPI.Controllers {
	public class LanguagesController : ApiController {
		private readonly Itbuser_LanguagesService _tbuser_LanguagesService;

		public LanguagesController(Itbuser_LanguagesService tbuser_LanguagesService) {
			_tbuser_LanguagesService = tbuser_LanguagesService;
		}

		public LanguagesController() {
		}

		public IEnumerable<tbuser_Languages_Model> Get() {
			return _tbuser_LanguagesService.List();
		}
	}
}