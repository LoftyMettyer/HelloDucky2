using System;
using OpenHRNexus.Common.Models;

namespace OpenHRNexus.Service.Interfaces
{
	public interface IBusinessProcessService
	{
//		IEnumerable<DynamicDataModel> GetData(int id);
//		IEnumerable<DynamicDataModel> GetData();
        WebFormModel GetWebFormForProcessAndUser(int processId, Guid userId);

    }
}
