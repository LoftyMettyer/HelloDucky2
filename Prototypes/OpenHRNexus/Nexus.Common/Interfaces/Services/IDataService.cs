using System;
using Nexus.Common.Models;
using Nexus.Common.Classes;
using System.Collections.Generic;

namespace Nexus.Common.Interfaces.Services
{
	public interface IDataService {

		WebFormModel InstantiateProcess(int processId, Guid userId, string language);

        ProcessStepResponse SubmitStepForUser(Guid stepId, Guid userId, WebFormModel form);
        IEnumerable<CalendarEventModel> GetReportData(int reportID, IEnumerable<IReportDataFilter> filters);
    }
}
