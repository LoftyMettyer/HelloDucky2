using System.Collections.Generic;
using OpenHRNexus.Common.Models;
using System;

namespace OpenHRNexus.Repository.Interfaces
{
	public interface IDataRepository
	{
		IEnumerable<DynamicDataModel> GetData(int id);
		IEnumerable<DynamicDataModel> GetData();
        WebForm GetWebForm(int id);
        WebFormModel PopulateFormWithData(WebForm webForm, Guid userId);

    }
}
