﻿using Microsoft.AspNet.Identity;
using Nexus.Common.Classes;
using Nexus.Common.Enums;
using Nexus.Common.Interfaces.Services;
using Nexus.Common.Models;
using Nexus.Sql_Repository.DatabaseClasses.Data;
using System;
using System.Collections.Generic;
using System.Security.Claims;
using System.Web;
using System.Web.Http;

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

        /// <summary>
        /// Controller constructor for use with Ninject
        /// </summary>
        /// <param name="dataService"></param>
        public ProcessController(IDataService dataService)
        {
            _dataService = dataService;
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
        [Authorize(Roles="OpenHRUser")]
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
        [Route("process/{processId:int, stepId: Guid}")]
        public IEnumerable<WebFormModel> GetProcessStep([FromUri] int processId, [FromUri] Guid? stepId)
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
        /// <param name="form"></param>
        /// <returns></returns>
        [Authorize(Roles = "OpenHRUser")]
        [Route("process")]
        public ProcessStepResponse PostProcessStep(WebFormModel form)
        {

            // Put some clever code in an attribute extension to validate that there is a identity getuserguid?
            // Maybe this is already covered by the authorize roles = OpenHRUser?
            var userId = new Guid(_identity.GetUserId());

            return _dataService.SubmitStepForUser(form.stepid, userId, form);
        }


    }
}