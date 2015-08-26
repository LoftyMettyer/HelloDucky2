using System;
using System.Collections.Generic;
using Nexus.Common.Models;

namespace Nexus.Common.Interfaces.Repository {
	public interface IDataRepository {
		IEnumerable<DynamicDataModel> GetData(int id);
		IEnumerable<DynamicDataModel> GetData();
		WebForm GetWebForm(int id);
		WebFormModel PopulateFormWithData(WebForm webForm, Guid userId);
	}
}
