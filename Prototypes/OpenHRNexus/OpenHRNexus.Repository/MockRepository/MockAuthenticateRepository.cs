using System;
using OpenHRNexus.Repository.Interfaces;
using OpenHRNexus.Repository.Messages;
using Repository.Enums;

namespace OpenHRNexus.Repository.MockRepository {
	public class MockAuthenticateRepository : IAuthenticateRepository {
		public RegisterNewUserMessage RequestAccount(string email)
		{
			return new RegisterNewUserMessage {
				UserID = Guid.NewGuid(), 
				Status = NewUserStatus.Success
			};
		}
	}
}