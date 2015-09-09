using System;
using Nexus.Common.Interfaces.Repository;
using Nexus.Common.Models;
using Nexus.Common.Classes;
using System.Collections.Generic;
using OpenHRNexus.Common.Enums;
using Nexus.Common.Interfaces;
using Nexus.Common.Interfaces.Services;

namespace Nexus.Service.Services {
	public class DataService : IDataService {
		private IDataRepository _dataRepository;

		public DataService(IDataRepository dataRepository) {
			_dataRepository = dataRepository;
		}

        public IEnumerable<CalendarEventModel> GetReportData(int reportID, IEnumerable<IReportDataFilter> filters)
        {
            // Some preprocesslogic

            // get definitions of the data
            return _dataRepository.GetReportData(reportID, filters);

        }

        WebFormModel IDataService.GetWebForm(int ProcessId, Guid userId, string language)
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

            WebForm webForm = _dataRepository.GetWebForm(ProcessId, language);
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

        ProcessStepResponse IDataService.SubmitStepForUser(Guid stepId, Guid userID, WebFormModel form)
        {
            var result = new ProcessStepResponse();

            // Find out what our next steps are.
            var currentStep = _dataRepository.GetProcessStep(stepId);

            currentStep.Validate();

            // Apply and security to fields submitted, i.e. if they've hacked values in, only allow through what they actually have access to.

            // Is this handled in the repository?
            //          _dataRepository.GetWriteableColumnsForUser();
            //            _dataRepository.GetWriteableTablesForUser();

            //_dataRepository.SetUserPermissions(); ??

            //            var currentForm = _dataRepository.GetWebForm(businessProcessId);

            // if its a save for later, well just do it!
            // Find out what the next step is.

            var nextStep = _dataRepository.GetProcessNextStep(currentStep);

            // Oooh they decided to save for later.

            // Commit the step (change later to put in below logic
            result = _dataRepository.CommitStep(stepId, userID, form);


            // Oooh they want to send an email
            switch (nextStep.Type)
            {
                case ProcessStepType.Email:
                    var emailService = new EmailService();
                    var processStepEmail = (ProcessStepEmail)nextStep;

                    //Todo: Determine what type of email template we are using so we can do some extra processing on the template, such as replacing certain placeholders, etc.

                    result = emailService.Send(processStepEmail.To, processStepEmail.Subject,
												string.Format(processStepEmail.Message, "Debbie Avery", "two day", "19/09/2015", "25/09/2015", "Holiday","Sorry it's short notice!", "http://www.bbc.co.uk", "http://www.bbc.co.uk", "http://www.bbc.co.uk", "http://www.bbc.co.uk"));
                    break;

                case ProcessStepType.StoredData:
                    //result = _dataRepository.SaveStepForLater(stepId, userID, form);
                    result = _dataRepository.CommitStep(stepId, userID, form);
                    break;

                default:
                    Console.WriteLine("Default case");
                    break;
            }
        

            return result;

        }


    }
}
