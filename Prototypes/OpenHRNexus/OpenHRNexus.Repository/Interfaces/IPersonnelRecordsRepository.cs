using System.Collections.Generic;
using OpenHRNexus.Repository;

namespace OpenHRNexus.Repository.Interfaces {
	public interface IPersonnelRecordsRepository {
		List<Personnel_Records> List();
	}
}
