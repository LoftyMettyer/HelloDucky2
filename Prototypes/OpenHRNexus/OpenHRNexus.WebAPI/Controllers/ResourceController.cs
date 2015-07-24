using System.Collections.Generic;
using System.Linq;
using System.Web.Http;

namespace OpenHRNexus.WebAPI.Controllers {
	public class ResourceController : ApiController {
		[HttpGet]
		public IEnumerable<KeyValuePair<string, string>> GetResourceValues([FromUri] List<string> id) {
			return id.Select(s => new KeyValuePair<string, string>(s, Resources.Resource.ResourceManager.GetString(s))).ToList();
		}
	}
}
