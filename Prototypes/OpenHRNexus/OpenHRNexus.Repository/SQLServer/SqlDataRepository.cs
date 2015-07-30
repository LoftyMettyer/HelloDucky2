using System.Collections.Generic;
using System.Data.Entity;
using System.Linq;
using OpenHRNexus.Common.Models;
using OpenHRNexus.Repository.Interfaces;

namespace OpenHRNexus.Repository.SQLServer
{
	public class SqlDataRepository : DbContext, IDataRepository
	{
		public IEnumerable<DynamicDataModel> GetData(int id)
		{
			var result = Data
				.Where(c => c.Id == id);

			return result.ToList();
		
		}

		public IEnumerable<DynamicDataModel> GetData()
		{
			return Data.ToList();
		}

		public virtual DbSet<DynamicDataModel> Data { get; set; }

	}
}
