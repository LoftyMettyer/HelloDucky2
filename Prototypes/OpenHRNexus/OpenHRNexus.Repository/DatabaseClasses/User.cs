using System;
using OpenHRNexus.Common.Interfaces;

namespace OpenHRNexus.Repository.DatabaseClasses {
	public class User : INexusUser {
		public Guid Id { get; set; }
		public string Role { get; set; }
		public DateTime LastConnectDateTime { get; set; }
		public int RecordId { get; set; }
	}
}