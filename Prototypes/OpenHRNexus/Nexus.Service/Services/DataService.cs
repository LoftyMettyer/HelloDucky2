using System;
using Nexus.Common.Interfaces.Repository;
using Nexus.Common.Models;
using Nexus.Service.Interfaces;
using Nexus.Common.Classes;
using System.Collections.Generic;
using Nexus.Common.Enums;

namespace Nexus.Service.Services {
	public class DataService : IDataService {
		private IDataRepository _dataRepository;

		public DataService(IDataRepository dataRepository) {
			_dataRepository = dataRepository;
		}

        WebFormModel IDataService.GetWebForm(int businessProcessId, Guid userId)
        {

            // Move to a factory for flexibility and eaiser reading?
       //     var businessProcess = _dataRepository.GetBusinessProcess(businessProcessId);

    //        if (businessProcess == null) return null;

            //var businessProcess = (BusinessProcessModel)_dataRepository.GetBusinessProcess(businessProcessId);

            //   BusinessProcessModel businessProcess2 = (BusinessProcessModel)_dataRepository.GetBusinessProcess(businessProcessId);
            //    var model = new BusinessProcessModel(businessProcess);
            //            var businessProcess = new BusinessProcessModel(_dataRepository);

            //, businessProcessId);
            //var webForm = businessProcess.GetFirstStep;
         


         //   firstStep.Translate("en-GB");


            WebForm webForm = _dataRepository.GetWebForm(businessProcessId);
       //     webForm.Translate("en-GB");


            var result = _dataRepository.PopulateFormWithData(webForm, userId);
            //var result = new WebFormModel();

      //      var result2 = _dataRepository.PopulateFormWithNavigationControls(webForm, userId);


            // Implement translation as a design pattern (a template one? - I can't remember - need to review training notes)
            //result.translate(language)

            // TODO - Fettle to get rid of recursive webform references. Ultimate solution is to return a different webform item to the internal
            // service and repository objects.

            foreach (var formField in result.fields)
            {
                formField.WebForm = null;
            }

            foreach (var formField in result.buttons)
            {
                formField.targeturl = string.Format(formField.targeturl, Guid.NewGuid());
                formField.WebForm = null;
            }

            //            result.form_fields.Remove[0];





            return result;
        }

        BusinessProcessStepResponse IDataService.SubmitStepForUser(Guid stepId, Guid userID, IEnumerable<KeyValuePair<int, string>> data)
        {

            // Get the form for this submitted step

            // Apply and security to fields submitted, i.e. if they've hacked values in, only allow through what they actually have access to.

            // Is this handled in the repository?
            //          _dataRepository.GetWriteableColumnsForUser();
            //            _dataRepository.GetWriteableTablesForUser();

            //_dataRepository.SetUserPermissions(); ??

//            var currentForm = _dataRepository.GetWebForm(businessProcessId);



            return new BusinessProcessStepResponse
            {
                Status = BusinessProcessStepStatus.Success,
                Message = "Success",
                FollowOnUrl = String.Empty
            };

        }


    }
}
