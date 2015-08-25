using System;
using System.Collections.Generic;
using System.Web.Http;
using Microsoft.AspNet.Identity;
using Nexus.Repository.Messages;
using Nexus.Service.Interfaces;

namespace Nexus.WebAPI.Controllers {
	public class AuthenticateController : ApiController {
		private readonly IAuthenticateService _authenticateService;

		public AuthenticateController() {
		}

		public AuthenticateController(IAuthenticateService authenticateService) {
			_authenticateService = authenticateService;
		}

		// GET api/authenticate/authenticate?parameter=email
		//todo: secure this controller actions so only authservice can access it.
		[HttpGet]
		public RegisterNewUserMessage Authenticate(string email, string userId) {
			return _authenticateService.RequestAccount(email, userId);
		}

		[HttpGet]
		public IEnumerable<string> GetClaims(string userId) {
			return _authenticateService.GetClaims(new Guid(userId));
		}

	}
}
