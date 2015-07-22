using System;

namespace OpenHRNexus.Interfaces.Common {
	public interface INexusUser {
		Guid Id { get; set; }
		string Role { get; set; }
		DateTime LastConnectDateTime { get; set; }
	}
}
