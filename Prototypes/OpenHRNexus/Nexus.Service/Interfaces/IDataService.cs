using System;
using Nexus.Common.Models;
using Nexus.Common.Classes;
using System.Collections.Generic;

namespace Nexus.Service.Interfaces
{
	public interface IDataService
	{
		//IEnumerable<DynamicDataModel> GetData(int id);
		//IEnumerable<DynamicDataModel> GetData();
        WebFormModel GetWebForm(int id, Guid userId);
        BusinessProcessStepResponse SubmitStepForUser(Guid stepId, Guid UserId, WebFormModel form);

    }
}
