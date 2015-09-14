using System;
using Nexus.Common.Models;
using Nexus.Common.Classes;
using System.Collections.Generic;

namespace Nexus.Common.Interfaces.Repository {
	public interface IDataRepository {
//		IEnumerable<DynamicDataModel> GetData(int id);
//		IEnumerable<DynamicDataModel> GetData();
		WebForm GetWebForm(int id, string language);
		WebFormModel PopulateFormWithData(WebForm webForm, Guid userId);
        WebFormModel PopulateFormWithNavigationControls(WebForm webForm, Guid userId);
        Process GetProcess(int Id);
        ProcessStepResponse SaveStepForLater(Guid stepId, Guid userID, WebFormModel form);
        ProcessStepResponse CommitStep(Guid stepId, Guid userID, WebFormModel form);
        IProcessStep GetProcessStep(Guid stepId);
        IProcessStep GetProcessNextStep(IProcessStep currentStep);
        IEnumerable<CalendarEventModel> GetReportData(int reportID, IEnumerable<IReportDataFilter> filters);

        Guid RecordProcessStep(WebForm form);

    }
}
