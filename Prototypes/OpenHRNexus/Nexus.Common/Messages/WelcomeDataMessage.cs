using System;

namespace Nexus.Common.Messages {
	public class WelcomeDataMessage {
		public Guid UserId { get; set; }
		public string Language { get; set; }
		public string Message { get; set; }
		public DateTime LastLoggedOn { get; set; }

		public string SecurityGroup {
			get { return "NotYetImplemented"; }
		}
	}
}
