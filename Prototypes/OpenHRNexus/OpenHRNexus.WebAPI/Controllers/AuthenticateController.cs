using System;
using System.Collections.Generic;
using System.Web.Http;
using Microsoft.AspNet.Identity;
using OpenHRNexus.Repository.Messages;
using OpenHRNexus.Service.Interfaces;

namespace OpenHRNexus.WebAPI.Controllers {	
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
		public RegisterNewUserMessage Authenticate(string parameter)
		{
			return _authenticateService.RequestAccount(parameter);
		}

		[HttpGet]
		public IEnumerable<string> GetClaims()
		{
		    string userId = User.Identity.GetUserId();

			return _authenticateService.GetRoles(new Guid(userId));
		}

	}
}
