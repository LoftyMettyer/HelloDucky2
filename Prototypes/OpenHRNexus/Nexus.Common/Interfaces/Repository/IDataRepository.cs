using System;
using Nexus.Common.Models;
using Nexus.Common.Classes;

namespace Nexus.Common.Interfaces.Repository {
	public interface IDataRepository {
//		IEnumerable<DynamicDataModel> GetData(int id);
//		IEnumerable<DynamicDataModel> GetData();
		WebForm GetWebForm(int id);
		WebFormModel PopulateFormWithData(WebForm webForm, Guid userId);
        WebFormModel PopulateFormWithNavigationControls(WebForm webForm, Guid userId);
        BusinessProcess GetBusinessProcess(int Id);
    }
}
