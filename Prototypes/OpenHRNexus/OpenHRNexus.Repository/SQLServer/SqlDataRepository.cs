using System.Collections.Generic;
using System.Data.Entity;
using OpenHRNexus.Common.Models;
using OpenHRNexus.Repository.Interfaces;

namespace OpenHRNexus.Repository.SQLServer
{
	public class SqlDataRepository : DbContext, IDataRepository
	{
		public IEnumerable<DynamicDataModel> GetData(int id)
		{
			return new List<DynamicDataModel>();
		}
	}
}
