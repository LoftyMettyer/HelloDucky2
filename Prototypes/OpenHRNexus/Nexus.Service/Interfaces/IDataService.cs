using System;
using Nexus.Common.Models;

namespace Nexus.Service.Interfaces
{
	public interface IDataService
	{
		//IEnumerable<DynamicDataModel> GetData(int id);
		//IEnumerable<DynamicDataModel> GetData();
        WebFormModel GetWebForm(int id, Guid userId);

    }
}
