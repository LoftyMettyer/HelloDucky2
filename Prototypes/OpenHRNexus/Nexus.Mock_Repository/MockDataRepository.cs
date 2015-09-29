using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Mail;
using Nexus.Common.Classes;
using Nexus.Common.Interfaces;
using Nexus.Common.Interfaces.Repository;
using Nexus.Common.Models;
using Nexus.Sql_Repository.DatabaseClasses.Data;
using System.Collections;
using System.Threading.Tasks;
using Nexus.Common.Enums;

namespace Nexus.Mock_Repository {
	public class MockProcessRepository : IProcessRepository {

		public ProcessFormElement GetWebForm(int id) {
			throw new NotImplementedException();
		}

		public ProcessFormElement PopulateFormWithData(ProcessFormElement webForm, Guid userId) {
			throw new NotImplementedException();
		}

        public Process GetProcess(int Id)
        {
            throw new NotImplementedException();
        }

        public ProcessStepResponse SaveStepForLater(Guid stepId, Guid userID, WebFormModel form)
        {
            throw new NotImplementedException();
        }

        public IProcessStep GetProcessStep(Guid stepId)
        {
            throw new NotImplementedException();
        }

        public IProcessStep GetProcessNextStep(IProcessStep currentStep)
        {
            throw new NotImplementedException();
        }

        public ProcessStepResponse CommitStep(Guid stepId, Guid userID, WebFormModel form)
        {
            throw new NotImplementedException();
        }

        public IEnumerable<CalendarEventModel> GetReportData(int reportID, IEnumerable<IReportDataFilter> filters)
        {
            throw new NotImplementedException();
        }

        public Guid RecordProcessStepForUser(ProcessFormElement form, Guid userId)
        {
            throw new NotImplementedException();
        }

        public IEnumerable<ProcessInFlow> GetProcesses(Guid userId)
        {
            throw new NotImplementedException();
        }

        public async Task<IEnumerable> GetData(int dataSourceId, IEnumerable<IReportDataFilter> filters)
        {
            throw new NotImplementedException();
        }

        public Type GetDataDefinition(int dataSourceId)
        {
            throw new NotImplementedException();
        }

        public ProcessEmailTemplate GetEmailTemplate(int id)
        {
            throw new NotImplementedException();
        }

        public MailMessage PopulateEmailWithData(IProcessStep step, Guid userId, ProcessEmailTemplate template)
        {
            throw new NotImplementedException();
        }

        public ProcessStepStatus RecordProcessStepForUser(WebFormDataModel formData, Guid userId)
        {
            throw new NotImplementedException();
        }

        public ProcessStepStatus RecordProcessStepForUser(Process process, WebFormDataModel formData, Guid userId)
        {
            throw new NotImplementedException();
        }

        public WebFormDataModel UpdateProcessWithUserVariables(Process process, WebFormDataModel formData, Guid userId)
        {
            throw new NotImplementedException();
        }

        public Process GetProcessForStep(Guid stepId)
        {
            throw new NotImplementedException();
        }
    }
}
