using System;
using Repository.Enums;

namespace OpenHRNexus.Repository.Messages
{
	public class RegisterNewUserMessage
	{
		public NewUserStatus Status { get; set; }
	}
}
