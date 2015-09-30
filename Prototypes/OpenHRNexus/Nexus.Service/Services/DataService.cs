using System;
using Nexus.Common.Interfaces.Repository;
using Nexus.Common.Models;
using Nexus.Common.Classes;
using System.Collections.Generic;
using OpenHRNexus.Common.Enums;
using Nexus.Common.Interfaces.Services;
using Nexus.Common.Enums;
using Nexus.Sql_Repository.DatabaseClasses.Data;
using System.Linq;
using Nexus.WebAPI.Handlers;
using System.Net.Mail;
using System.Collections;
using System.Threading.Tasks;
using Nexus.Common.Interfaces;


namespace Nexus.Service.Services
{
    public class DataService : IDataService
    {
		private IProcessRepository _dataRepository;
        private ITranslation _translation;
        private string _callingURL;
        private string _authenticationServiceURL;

        public string CallingURL
        {
            get
            { return _callingURL; }
            set { _callingURL = value; }
        }

        public string AuthenticationServiceURL
        {
            get
            { return _authenticationServiceURL; }
            set { _authenticationServiceURL = value; }
        }

        public DataService(IProcessRepository dataRepository, ITranslation translation)
        {
			_dataRepository = dataRepository;
            _translation = translation;
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


        public IEnumerable<SummaryDataModel> GetSummaryData(Guid userId, int reportID, IEnumerable<IReportDataFilter> filters)
        {
            var rawData = _dataRepository.GetProcesses(userId);

            return rawData
                .GroupBy(c => c.ProcessName)
                .Select(grp => new SummaryDataModel() { category = grp.Key, value = (int)grp.Count() })
                .ToList();

        }

        public async Task<IEnumerable> GetData(int dataSourceId, IEnumerable<IReportDataFilter> filters)
        {
            var data = await _dataRepository.GetData(dataSourceId, filters);
            return data;
        }


        public Type GetDataDefinition(int dataSourceId)
        {
            return _dataRepository.GetDataDefinition(dataSourceId);
        }

        WebFormModel IDataService.InstantiateProcess(int ProcessId, Guid userId, string language)
        {

            var process = _dataRepository.GetProcess(ProcessId);

            var webForm = process.GetEntryPoint();


            // Yes , we could use the webform above (and very soon we will), but for the moment the translation isn't
            // quite hooked in properly.
            // ProcessFormElement webForm = _dataRepository.GetWebForm(firstStep.id);

            //var stepId = _dataRepository.RecordProcessStepForUser(webForm, userId);




            _translation.Language = language;

            var populatedForm = _dataRepository.PopulateFormWithData(webForm, userId);
            populatedForm.Translate(_translation);
            populatedForm.SetButtonEndpoints(_callingURL, Guid.NewGuid());


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
                fields = form.Fields.OrderBy(s => s.sequence).ToList(),
                buttons = form.Buttons.Select(b => new WebFormButtonModel
                {
                    title = b.Title,
                    TargetUrl = b.TargetUrl,
                    action = b.Action
                }).ToList()
            };


            // TODO - Fettle to get rid of recursive webform references. Ultimate solution is to return a different webform item to the internal
            // service and repository objects.

            foreach (var formField in result.fields)
            {
                formField.WebForm = null;
            }

            //foreach (var formField in result.buttons)
            //{
            //    formField.WebForm = null;
            //}

            return result;

        }


        async Task<ProcessStepResponse> IDataService.SubmitStepForUser(Guid stepId, Guid userID, WebFormDataModel formData)
        {
            var result = new ProcessStepResponse();


            WebFormModel form = new WebFormModel();

            // Conversion to internal data mappings, to be reworked once endpoints are finalised.
            form.stepid = formData.stepid;
            form.fields = new List<WebFormField>();
            foreach (var field in formData.data)
            {
                form.fields.Add(new WebFormField() { elementid = field.Key, value = field.Value.ToString() });
            }

            // Calculate which process we're running.

            var process = _dataRepository.GetProcessForStep(stepId);

            var currentProcessVariables = _dataRepository.UpdateProcessWithUserVariables(process, formData, userID);

            // Merge variable list with default values (For demo only - requires move to instaniate process and merge with above!)
            currentProcessVariables.data.Add("we_1_1", "Nick");
            currentProcessVariables.data.Add("we_2_2", "Gibson");


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

            // currently hardcoded to always return an email as the next step
            var nextStep = _dataRepository.GetProcessNextStep(currentStep);

            // Oooh they decided to save for later.

            // Commit the step (change later to put in below logic
            result = _dataRepository.CommitStep(stepId, userID, form);


            // Oooh they want to send an email
            switch (nextStep.Type)
            {
                case ProcessElementType.Email:

                    var emailTemplate = _dataRepository.GetEmailTemplate(1);

                    emailTemplate._translation = _translation;
                    emailTemplate.Destinations = emailTemplate.Destinations;        // May be calculated later?
                    emailTemplate.Variables = currentProcessVariables.data;

                    // Add authenication to each follow on action
                    foreach (var button in emailTemplate.FollowOnActions)
                    {
                        button.TargetUrl = "";
                    }

                    foreach (var button in emailTemplate.FollowOnActions)
                    {
                        var authenticationCode = await AuthenticationServiceHandler.GetUserToken(_authenticationServiceURL, userID, button.TargetStep);
                        button.TargetUrl = string.Format("{0}UI/home/postprocessstep?userid={1}&code={3}&purpose={2}"
                            , _callingURL, userID, button.TargetStep
                            , authenticationCode);
                    }

                    var emailService = new EmailService();
                    MailMessage message = emailTemplate.GenerateMailMessage();
                    result = emailService.Send(message);

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
