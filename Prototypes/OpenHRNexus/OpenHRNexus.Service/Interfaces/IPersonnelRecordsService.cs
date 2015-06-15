using System.Collections.Generic;
using OpenHRNexus.Common.Models;

namespace OpenHRNexus.Service.Interfaces {
	public interface IPersonnelRecordsService {
		List<Personnel_Records_Model> List();
	}
}
