﻿using System;
using Nexus.Common.Models;
using Nexus.Common.Classes;
using System.Collections.Generic;
using Nexus.Sql_Repository.DatabaseClasses.Data;
using Nexus.Common.Enums;
using System.Collections;
using System.Threading.Tasks;

namespace Nexus.Common.Interfaces.Services
{
	public interface IDataService {

        string CallingURL { get; set; }
        string AuthenticationServiceURL { get; set; }

        WebFormModel InstantiateProcess(int processId, Guid userId, string language);

        Task<ProcessStepResponse> SubmitStepForUser(Guid stepId, Guid userId, WebFormDataModel formData);
        IEnumerable<CalendarEventModel> GetReportData(int reportID, IEnumerable<IReportDataFilter> filters);
        Task<IEnumerable> GetData(int dataSourceId, IEnumerable<IReportDataFilter> filters);
        Type GetDataDefinition(int dataSourceId);
        IEnumerable<ProcessInFlow> GetEntitiesForUser(EntityType type, Guid userId);

    }
}
