using System.Collections.Generic;
using OpenHRNexus.Common.Models;
using OpenHRNexus.Repository.Interfaces;

namespace OpenHRNexus.Repository.MockRepository
{
	public class MockDataRepository : IDataRepository
	{
		public IEnumerable<DynamicDataModel> GetData()
		{

			var data = new List<DynamicDataModel>
			{
				new DynamicDataModel
				{
					Id = 1,
					Column1 = "Jack",
					Column2 = "Jones",
					Column3 = "dob",
					Column4 = "number4"
				},
				new DynamicDataModel
				{
					Id = 1,
					Column1 = "Fred",
					Column2 = "Smith",
					Column3 = "12/8/1975",
					Column4 = "unknown field"
				},
			};

			return data;
		}

	}
}
