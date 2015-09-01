using System;

namespace Nexus.Common.Interfaces {
	public interface INexusUser {
		Guid Id { get; set; }
	//	string Role { get; set; }
	//	DateTime LastConnectDateTime { get; set; }
		int RecordId { get; set; }
	}
}
