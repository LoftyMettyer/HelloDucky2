using System.Web.Http;
using OpenHRNexus.Repository.Messages;
using OpenHRNexus.Service.Interfaces;

namespace OpenHRNexus.WebAPI.Controllers {
	[Authorize(Roles = "Admin")]
	public class AuthenticateController : ApiController {
		private readonly IAuthenticateService _authenticateService;

		public AuthenticateController() {
		}

		public AuthenticateController(IAuthenticateService authenticateService) {
			_authenticateService = authenticateService;
		}

		// GET api/authenticate/authenticate?parameter=email
		[HttpGet]
		public RegisterNewUserMessage Authenticate(string parameter)
		{
			return _authenticateService.RequestAccount(parameter);
		}
	}
}
