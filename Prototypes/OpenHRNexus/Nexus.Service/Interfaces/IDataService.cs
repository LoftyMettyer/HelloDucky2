﻿using System;
using System.Threading.Tasks;
using Nexus.Common.Models;
using Nexus.Common.Classes;

namespace Nexus.Service.Interfaces {
	public interface IDataService {
		//IEnumerable<DynamicDataModel> GetData(int id);
		//IEnumerable<DynamicDataModel> GetData();
		WebFormModel GetWebForm(int id, Guid userId);
		BusinessProcessStepResponse SubmitStepForUser(Guid stepId, Guid userId, WebFormModel form);
	}
}