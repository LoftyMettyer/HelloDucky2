using System;
using System.Collections.Generic;
using System.Linq;
using OpenHRNexus.Repository.Interfaces;

namespace OpenHRNexus.Repository.SQLServer
{
	public class SqlUserRepository  : IWelcomeMessageData
	{
		private Guid _userID;
		private string _language;

		public SqlUserRepository(Guid userID, string language)
		{
			_userID = userID;
			_language = language;
		}

		public Guid UserID
		{
			set { _userID = value; }
		}

		public string Language
		{
			set { _language = value; }
		}

		public string WelcomeMessageData
		{
			get { return "Mrs Debbie Avery"; }
		}

		public DateTime LastLoginDateTime
		{
			get { return new DateTime(1975,08,10); }
		}

		public string SecurityGroup
		{
			get { return "Unknown Group"; }
		}
	}
}
