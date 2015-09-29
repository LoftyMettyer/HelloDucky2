using System;
using Nexus.Common.Models;
using Nexus.Common.Classes;
using System.Collections.Generic;
using Nexus.Sql_Repository.DatabaseClasses.Data;
using System.Net.Mail;
using System.Linq;
using System.Collections;
using System.Threading.Tasks;
using Nexus.Common.Enums;

namespace Nexus.Common.Interfaces.Repository {
	public interface IProcessRepository {

		//ProcessFormElement GetWebForm(int id);
        ProcessFormElement PopulateFormWithData(ProcessFormElement webForm, Guid userId);
        ProcessEmailTemplate GetEmailTemplate(int id);
        Process GetProcess(int Id);
        ProcessStepResponse SaveStepForLater(Guid stepId, Guid userID, WebFormModel form);
        ProcessStepResponse CommitStep(Guid stepId, Guid userID, WebFormModel form);
        IProcessStep GetProcessStep(Guid stepId);
        IProcessStep GetProcessNextStep(IProcessStep currentStep);
        WebFormDataModel UpdateProcessWithUserVariables(Process process, WebFormDataModel formData, Guid userId);

        IEnumerable<ProcessInFlow> GetProcesses(Guid userId);

        IEnumerable<CalendarEventModel> GetReportData(int reportID, IEnumerable<IReportDataFilter> filters);
        Task<IEnumerable> GetData(int dataSourceId, IEnumerable<IReportDataFilter> filters);
        Type GetDataDefinition(int dataSourceId);

    }
}
