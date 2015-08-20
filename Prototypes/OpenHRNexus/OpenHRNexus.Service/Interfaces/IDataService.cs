﻿using System;
using System.Collections.Generic;
using OpenHRNexus.Common.Models;

namespace OpenHRNexus.Service.Interfaces
{
	public interface IDataService
	{
		IEnumerable<DynamicDataModel> GetData(int id);
		IEnumerable<DynamicDataModel> GetData();
        WebForm GetWebForm(int id);

    }
}
