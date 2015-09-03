using System;
using Nexus.Common.Models;
using Nexus.Common.Classes;
using System.Collections.Generic;

namespace Nexus.Common.Interfaces.Repository {
	public interface IDataRepository {
//		IEnumerable<DynamicDataModel> GetData(int id);
//		IEnumerable<DynamicDataModel> GetData();
		WebForm GetWebForm(int id);
		WebFormModel PopulateFormWithData(WebForm webForm, Guid userId);
        WebFormModel PopulateFormWithNavigationControls(WebForm webForm, Guid userId);
        BusinessProcess GetBusinessProcess(int Id);
        BusinessProcessStepResponse SaveStepForLater(Guid stepId, Guid userID, WebFormModel form);
        BusinessProcessStepResponse CommitStep(Guid stepId, Guid userID, WebFormModel form);
        IBusinessProcessStep GetBusinessProcessStep(Guid stepId);
        IBusinessProcessStep GetBusinessProcessNextStep(IBusinessProcessStep currentStep);
        IEnumerable<CalendarEventModel> GetReportData(int reportID, IEnumerable<IReportDataFilter> filters);

    }
}
