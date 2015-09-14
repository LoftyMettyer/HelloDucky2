using System;
using Nexus.Common.Models;
using Nexus.Common.Classes;
using System.Collections.Generic;
using Nexus.Sql_Repository.DatabaseClasses.Data;

namespace Nexus.Common.Interfaces.Repository {
	public interface IProcessRepository {

		ProcessFormElement GetWebForm(int id, string language);
        ProcessFormElement PopulateFormWithData(ProcessFormElement webForm, Guid userId);


    //    WebFormModel PopulateFormWithNavigationControls(WebForm webForm, Guid userId);
        Process GetProcess(int Id);
        ProcessStepResponse SaveStepForLater(Guid stepId, Guid userID, WebFormModel form);
        ProcessStepResponse CommitStep(Guid stepId, Guid userID, WebFormModel form);
        IProcessStep GetProcessStep(Guid stepId);
        IProcessStep GetProcessNextStep(IProcessStep currentStep);
        IEnumerable<CalendarEventModel> GetReportData(int reportID, IEnumerable<IReportDataFilter> filters);

        Guid RecordProcessStepForUser(ProcessFormElement form, Guid userId);

        IEnumerable<ProcessInFlow> GetProcesses(Guid userId);
    }
}
