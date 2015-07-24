using System;
using OpenHRNexus.Repository.Interfaces;

namespace OpenHRNexus.Repository.SQLServer {
	public class SqlUserRepository : IWelcomeMessageDataRepository {
		private Guid _userId;
		private string _language;

		public SqlUserRepository() {
		}

		public SqlUserRepository(Guid userId, string language) {
			_userId = userId;
			_language = language;
		}

		public Guid UserId {
			set { _userId = value; }
		}

		public string Language {
			set { _language = value; }
		}

		public string WelcomeMessageData {
			get { return "Mrs Debbie Avery"; }
		}

		public DateTime LastLoginDateTime {
			get { return new DateTime(1975, 08, 10); }
		}

		public string SecurityGroup {
			get { return "Unknown Group"; }
		}
	}
}
