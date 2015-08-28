using System;
using System.Collections.Generic;
using System.Linq;
using Nexus.Common.Classes;
using Nexus.Common.Interfaces.Repository;
using Nexus.Common.Models;

namespace Nexus.Mock_Repository {
	public class MockDataRepository : IDataRepository {
		public IEnumerable<DynamicDataModel> GetData(int id) {
			var result = GetData().Where(m => m.Id == id);
			return result;
		}

		public IEnumerable<DynamicDataModel> GetData() {

            var data = new List<DynamicDataModel>();
			//{
			//	new DynamicDataModel
			//	{
			//		Id = 1,
			//		Column1 = "Jack",
			//		Column2 = "Jones",
			//		Column3 = "dob",
			//		Column4 = "number4"
			//	},
			//	new DynamicDataModel
			//	{
			//		Id = 1,
			//		Column1 = "Fred",
			//		Column2 = "Smith",
			//		Column3 = "12/8/1975",
			//		Column4 = "unknown field"
			//	},
			//};

			return data;

		}

		public WebForm GetWebForm(int id) {
			throw new NotImplementedException();
		}

		public WebFormModel PopulateFormWithData(WebForm webForm, Guid userId) {
			throw new NotImplementedException();
		}

        public BusinessProcess GetBusinessProcess(int Id)
        {
            throw new NotImplementedException();
        }
    }
}
