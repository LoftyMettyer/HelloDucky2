using System;
using Nexus.Common.Interfaces.Repository;
using Nexus.Common.Models;
using Nexus.Common.Classes;
using System.Collections.Generic;
using OpenHRNexus.Common.Enums;
using Nexus.Common.Interfaces;
using Nexus.Common.Interfaces.Services;
using Nexus.Common.Enums;
using Nexus.Sql_Repository.DatabaseClasses.Data;
using System.Linq;

namespace Nexus.Service.Services {
	public class DataService : IDataService {
		private IProcessRepository _dataRepository;

		public DataService(IProcessRepository dataRepository) {
			_dataRepository = dataRepository;
		}

        /// <summary>
        /// TODO - This function will return all entities types. At present all it returns is processes in flow.
        /// </summary>
        /// <param name="type"></param>
        /// <param name="userId"></param>
        /// <returns></returns>
        public IEnumerable<ProcessInFlow> GetEntitiesForUser(EntityType type, Guid userId)
        {
            return _dataRepository.GetProcesses(userId).ToList();
        }

        public IEnumerable<CalendarEventModel> GetReportData(int reportID, IEnumerable<IReportDataFilter> filters)
        {
            // Some preprocesslogic

            // get definitions of the data
            return _dataRepository.GetReportData(reportID, filters);

        }
 
        WebFormModel IDataService.InstantiateProcess(int ProcessId, Guid userId, string language)
        {

            var process = _dataRepository.GetProcess(ProcessId);

            var firstStep = process.GetEntryPoint();


            // Yes , we could use the webform above (and very soon we will), but for the moment the translation isn't
            // quite hooked in properly.
            ProcessFormElement webForm = _dataRepository.GetWebForm(firstStep.id, language);
            //     webForm.Translate("en-GB");

            var stepId = _dataRepository.RecordProcessStepForUser(webForm, userId);

            var populatedForm = _dataRepository.PopulateFormWithData(webForm, userId);


            populatedForm.SetButtonEndpoints(stepId);




            // Implement translation as a design pattern (a template one? - I can't remember - need to review training notes)
            //result.translate(language)

            // Tempry hack to convert internal to external webform models
            return PrepareWebFormModelFromInternalClass(populatedForm);


        }

        /// <summary>
        /// This is a tempry code stub to separate internal from external webformmodels. To be replaced by an interfaced object -
        /// but my powers of magic are required elsewhere. Hence marked function as obsolete
        /// </summary>
        /// <param name="form"></param>
        /// <returns></returns>
        [Obsolete("A temporary code stub while I figure out how to implement WebFormModel and ProcessFormElement to pull from the same interface.")]
        private WebFormModel PrepareWebFormModelFromInternalClass(ProcessFormElement form)
        {

            var result = new WebFormModel
            {
                id = form.id,
                stepid = Guid.NewGuid(),
                name = form.Name,
                fields = form.Fields,
                buttons = form.Buttons
            };

            // TODO - Fettle to get rid of recursive webform references. Ultimate solution is to return a different webform item to the internal
            // service and repository objects.

            foreach (var formField in result.fields)
            {
                formField.WebForm = null;
            }

            foreach (var formField in result.buttons)
            {
                formField.WebForm = null;
            }

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
                case ProcessElementType.Email:
                    var emailService = new EmailService();
                    var processStepEmail = (ProcessStepEmail)nextStep;

                    //Todo: Determine what type of email template we are using so we can do some extra processing on the template, such as replacing certain placeholders, etc.

                    result = emailService.Send(processStepEmail.To, processStepEmail.Subject,
												string.Format(processStepEmail.Message, "Debbie Avery", "two day", "19/09/2015", "25/09/2015", "Holiday","Sorry it's short notice!", "http://www.bbc.co.uk", "http://www.bbc.co.uk", "http://www.bbc.co.uk", "http://www.bbc.co.uk"));
                    break;

                case ProcessElementType.StoredData:
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
