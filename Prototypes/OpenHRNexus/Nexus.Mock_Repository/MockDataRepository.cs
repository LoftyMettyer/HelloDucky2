using System;
using System.Collections.Generic;
using Nexus.Common.Classes;
using Nexus.Common.Interfaces;
using Nexus.Common.Interfaces.Repository;
using Nexus.Common.Models;

namespace Nexus.Mock_Repository {
	public class MockDataRepository : IDataRepository {

		public WebForm GetWebForm(int id, string language) {
			throw new NotImplementedException();
		}

		public WebFormModel PopulateFormWithData(WebForm webForm, Guid userId) {
			throw new NotImplementedException();
		}

        public BusinessProcess GetBusinessProcess(int Id)
        {
            throw new NotImplementedException();
        }

        public WebFormModel PopulateFormWithNavigationControls(WebForm webForm, Guid userId)
        {
            throw new NotImplementedException();
        }

        public BusinessProcessStepResponse SaveStepForLater(Guid stepId, Guid userID, WebFormModel form)
        {
            throw new NotImplementedException();
        }

        public IBusinessProcessStep GetBusinessProcessStep(Guid stepId)
        {
            throw new NotImplementedException();
        }

        public IBusinessProcessStep GetBusinessProcessNextStep(IBusinessProcessStep currentStep)
        {
            throw new NotImplementedException();
        }

        public BusinessProcessStepResponse CommitStep(Guid stepId, Guid userID, WebFormModel form)
        {
            throw new NotImplementedException();
        }

        public IEnumerable<CalendarEventModel> GetReportData(int reportID, IEnumerable<IReportDataFilter> filters)
        {
            throw new NotImplementedException();
        }
    }
}
