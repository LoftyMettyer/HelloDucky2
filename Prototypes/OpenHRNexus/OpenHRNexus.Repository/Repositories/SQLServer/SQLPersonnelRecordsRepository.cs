using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OpenHRNexus.Repository.Interfaces;

namespace OpenHRNexus.Repository.Repositories.SQLServer {
	public class SQLPersonnelRecordsRepository : IPersonnelRecordsRepository {
		public List<Personnel_Records> List() {
			using (var db = new OpenHRNexusEntities()) {
				return db.Personnel_Records.ToList();
			}
		}
	}
}
