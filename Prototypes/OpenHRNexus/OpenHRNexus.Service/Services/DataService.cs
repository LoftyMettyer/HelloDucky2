using System.Collections.Generic;
using OpenHRNexus.Common.Models;
using OpenHRNexus.Repository.Interfaces;
using OpenHRNexus.Service.Interfaces;

namespace OpenHRNexus.Service.Services
{
	public class DataService : IDataService
	{
		private IDataRepository _dataRepository;

		public DataService(IDataRepository dataRepository) {
			_dataRepository = dataRepository;
		}

		public IEnumerable<DynamicDataModel> GetData(int id)
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
