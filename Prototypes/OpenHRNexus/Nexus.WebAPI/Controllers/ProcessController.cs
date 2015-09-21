using Microsoft.AspNet.Identity;
using Nexus.Common.Classes;
using Nexus.Common.Enums;
using Nexus.Common.Interfaces.Services;
using Nexus.Common.Models;
using Nexus.Sql_Repository.DatabaseClasses.Data;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Security.Claims;
using System.Threading.Tasks;
using System.Web;
using System.Web.Http;
using Nexus.WebAPI.Handlers;

namespace Nexus.WebAPI.Controllers
{
    /// <summary>
    /// Controller for Processes
    /// </summary>
    [RoutePrefix("api/process")]
    public class ProcessController : ApiController
    {

        private readonly IDataService _dataService;

        private ClaimsIdentity _identity;
        private string _language;

        private string GetApplicationSchemeName()
        {
            return HttpContext.Current.Request.Url.Scheme + "://" +
                  HttpContext.Current.Request.Url.Authority +
                  HttpContext.Current.Request.ApplicationPath.TrimEnd(Convert.ToChar("/")) + "/";

        }

        /// <summary>
        /// Controller constructor for use with Ninject
        /// </summary>
        /// <param name="dataService"></param>
        public ProcessController(IDataService dataService)
        {
            _dataService = dataService;
            _dataService.CallingURL = GetApplicationSchemeName();
            _identity = User.Identity as ClaimsIdentity;
            _language = HttpContext.Current.Request.UserLanguages[0].ToLowerInvariant().Trim();
        }

        /// <summary>
        /// Controller constructor with injection from Unit Test projects
        /// </summary>
        /// <param name="dataService"></param>
        /// <param name="claims"></param>
        /// <param name="language"></param>
        public ProcessController(IDataService dataService, ClaimsIdentity claims, string language)
        {
            _identity = claims;
            _dataService = dataService;
            _language = language;
        }



        /// <summary>
        /// Request a list of processes currently in mid-flow
        /// </summary>
        /// <returns>A JSON object containing all in-flow processes</returns>
        [Authorize(Roles = "OpenHRUser")]
        [Route("pendingprocesses")]
        public IEnumerable<ProcessInFlow> GetPendingProcesses()
        {
            var userId = new Guid(_identity.GetUserId());
            return _dataService.GetEntitiesForUser(EntityType.ProcessInFlow, userId);
        }

        /// <summary>
        /// Delete ALL processes currently in mid-flow
        /// </summary>
        /// <returns>Success flag</returns>
        [Authorize(Roles = "OpenHRUser")]
        [Route("pendingprocesses")]
        public IEnumerable<string> DeletePendingProcesses()
        {
            //Delete all pending processes
            return new[] { "Delete All" };
        }

        /// <summary>
        /// Delete a single process currently in mid-flow
        /// </summary>
        /// <param name="processId"></param>
        /// <returns>Success flag</returns>
        [Authorize(Roles = "OpenHRUser")]
        [Route("pendingprocesses/{processId:int}")]
        public IEnumerable<string> DeletePendingProcesses([FromUri] int processId)
        {
            //Delete specified pending process
            return new[] { "Delete " + processId };
        }

        /// <summary>
        /// Instantiates a business process.Returns a pre-populated, translated WebFormModel
        /// </summary>
        /// <param name="processId"></param>
        /// <param name="stepId"></param>
        /// <returns></returns>
        [Authorize(Roles = "OpenHRUser")]
        [Route("{processId:int}")]
        public IEnumerable<WebFormModel> GetProcessStep([FromUri] int processId)
        {

            // if not step id, start the process, else get an existing step

            var openHRDbGuid = new Guid(_identity.GetUserId());
            List<WebFormModel> form = new List<WebFormModel>();
            WebFormModel webForm;

            if (openHRDbGuid == null || openHRDbGuid == Guid.Empty)
            {
                // Berties error handler goes here ?
            }
            else
            {
                webForm = _dataService.InstantiateProcess(processId, openHRDbGuid, _language);
                form.Add(webForm);
            }

            IEnumerable<WebFormModel> webFormModels = form;
            return webFormModels;


        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="formData"></param>
        /// <returns></returns>
        [Authorize(Roles = "OpenHRUser")]
        [Route("")]
        public ProcessStepResponse PostProcessStep(WebFormDataModel formData)
        {

            // Put some clever code in an attribute extension to validate that there is a identity getuserguid?
            // Maybe this is already covered by the authorize roles = OpenHRUser?
            var userId = new Guid(_identity.GetUserId());

            return  _dataService.SubmitStepForUser(formData.stepid, userId, formData);

        }


        /// <summary>
        /// 
        /// </summary>
        /// <param name="formData">From Body (application/json)</param>
        /// <param name="userId"></param>
        /// <param name="code"></param>
        /// <returns></returns>
        [AllowAnonymous]
        [Route("")]
        public async Task<string> PostProcessStep([FromBody]WebFormDataModel formData, string userId, string code)
        {
            bool isValidToken = await AuthenticationServiceHandler.PostProcessStep(userId, code);

            if (isValidToken)
            {
                // Owasp data cleansing
                formData.DataCleanse();

                var result =  _dataService.SubmitStepForUser(formData.stepid, new Guid(userId), formData);

            }

            return isValidToken ? "OK" : "Not OK";

        }



    }
}
