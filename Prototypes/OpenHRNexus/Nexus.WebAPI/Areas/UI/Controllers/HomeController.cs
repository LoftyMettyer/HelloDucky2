using System;
using System.Threading.Tasks;
using System.Web.Http;
using System.Web.Mvc;
using Nexus.Common.Models;
using Nexus.WebAPI.Handlers;

namespace Nexus.WebAPI.Areas.UI.Controllers
{
    /// <summary>
    /// Controller for User Interaction only
    /// </summary>
    public class HomeController : Controller
    {
        // GET: UI/Home
        /// <summary>
        /// The home page for user interaction
        /// </summary>
        /// <returns></returns>
        public ActionResult Index()
        {            
            return View();
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="formData">Optional form body</param>
        /// <param name="userId">user GUID for this request</param>
        /// <param name="code">code provided to authenticate this request</param>
        /// <param name="purpose">purpose of this request</param>
        /// <returns></returns>
        public async Task<ActionResult> PostProcessStep([FromBody]WebFormDataModel formData, string userId, string code, string purpose)
        {
            bool isValidToken = await AuthenticationServiceHandler.PostProcessStep(userId, code, purpose);
          
            ViewBag.isValidToken = isValidToken ? "valid" : "not valid";
            
            return View("Index");

        }


    }
}