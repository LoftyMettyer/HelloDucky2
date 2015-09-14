using System;
using System.Collections.Generic;
using Nexus.Common.Classes;
using Nexus.Common.Interfaces;
using Nexus.Common.Interfaces.Repository;
using Nexus.Common.Models;
using Nexus.Sql_Repository.DatabaseClasses.Data;

namespace Nexus.Mock_Repository {
	public class MockProcessRepository : IProcessRepository {

		public WebForm GetWebForm(int id, string language) {
			throw new NotImplementedException();
		}

		public WebFormModel PopulateFormWithData(WebForm webForm, Guid userId) {
			throw new NotImplementedException();
		}

        public Process GetProcess(int Id)
        {
            throw new NotImplementedException();
        }

        public WebFormModel PopulateFormWithNavigationControls(WebForm webForm, Guid userId)
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

        public Guid RecordProcessStepForUser(WebForm form, Guid userId)
        {
            throw new NotImplementedException();
        }

        public IEnumerable<ProcessInFlow> GetProcesses(Guid userId)
        {
            throw new NotImplementedException();
        }
    }
}
