using Microsoft.AspNet.Identity;
using Nexus.Common.Enums;
using Nexus.Common.Interfaces.Services;
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

    }
}