﻿using System;
using System.Collections.Generic;
using OpenHRNexus.Repository.Interfaces;
using OpenHRNexus.Repository.Messages;
using OpenHRNexus.Service.Interfaces;

namespace OpenHRNexus.Service.Services {
	public class AuthenticateService : IAuthenticateService {
		private readonly IAuthenticateRepository _authenticateRepository;

		public AuthenticateService(IAuthenticateRepository auhenticateRepository) {
			_authenticateRepository = auhenticateRepository;
		}

		public RegisterNewUserMessage RequestAccount(string email, string userId) {
			return _authenticateRepository.RequestAccount(email, userId);
		}

		public IEnumerable<string> GetClaims(Guid userId) {
			return _authenticateRepository.GetUserPermissions(userId);
		}
	}
}