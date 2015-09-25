using System;
using Nexus.Common.Models;
using Nexus.Common.Classes;
using System.Collections.Generic;
using Nexus.Sql_Repository.DatabaseClasses.Data;
using Nexus.Common.Enums;
using System.Collections;

namespace Nexus.Common.Interfaces.Services
{
	public interface IDataService {

        string CallingURL { get; set; }
        string AuthenticationServiceURL { get; set; }

        WebFormModel InstantiateProcess(int processId, Guid userId, string language);

        ProcessStepResponse SubmitStepForUser(Guid stepId, Guid userId, WebFormDataModel formData);
        IEnumerable<CalendarEventModel> GetReportData(int reportID, IEnumerable<IReportDataFilter> filters);
        IEnumerable GetData(int dataSourceId, IEnumerable<IReportDataFilter> filters);
        IEnumerable<ProcessInFlow> GetEntitiesForUser(EntityType type, Guid userId);

    }
}
