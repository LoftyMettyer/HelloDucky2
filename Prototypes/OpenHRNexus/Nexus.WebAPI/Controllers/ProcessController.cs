using System.Collections.Generic;
using System.Web.Http;

namespace Nexus.WebAPI.Controllers
{
    /// <summary>
    /// Controller for Processes
    /// </summary>
    [RoutePrefix("api/process")]
    public class ProcessController : ApiController
    {
        /// <summary>
        /// Request a list of processes currently in mid-flow
        /// </summary>
        /// <returns>A JSON object containing all in-flow processes</returns>
        [Authorize(Roles="OpenHRUser")]
        [Route("pendingprocesses")]
        public IEnumerable<string> GetPendingProcesses()
        {
            //Get all pending processes
            return new[] {"OK"};
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