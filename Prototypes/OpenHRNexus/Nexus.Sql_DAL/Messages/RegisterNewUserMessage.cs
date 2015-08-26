using System;
using Repository.Enums;

namespace Nexus.Repository.Messages
{
	public class RegisterNewUserMessage
	{
		public NewUserStatus Status { get; set; }
	}
}
