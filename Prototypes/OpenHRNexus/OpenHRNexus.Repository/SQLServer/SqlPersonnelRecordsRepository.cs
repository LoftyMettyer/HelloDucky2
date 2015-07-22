using System.Collections.Generic;
using System.Linq;
using OpenHRNexus.Repository.Interfaces;

namespace OpenHRNexus.Repository.SQLServer {
	public class SqlPersonnelRecordsRepository : IPersonnelRecordsRepository {
		public List<Personnel_Records> List() {
			using (var db = new OpenHRNexusEntities()) {
				return db.Personnel_Records.ToList();
			}
		}
	}
}
