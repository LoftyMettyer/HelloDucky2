using System;
using System.Collections.Generic;
using System.Web.Http;
using Nexus.Common.Messages;
using Nexus.Common.Interfaces.Services;

namespace Nexus.WebAPI.Controllers {
	public class WelcomeController : ApiController {
		private readonly IWelcomeService _welcomeService;

		public WelcomeController() {
		}

		public WelcomeController(IWelcomeService welcomeService) {
			_welcomeService = welcomeService;
		}

		// GET api/authenticate/authenticate?parameter=email
		//todo: secure this controller actions so only authservice can access it.
		[HttpGet]
		public RegisterNewUserMessage Authenticate(string email, string userId) {
			return _welcomeService.RequestAccount(email, userId);
		}

		[HttpGet]
		public IEnumerable<string> GetClaims(string userId) {
			return _welcomeService.GetClaims(new Guid(userId));
		}

	}
}
