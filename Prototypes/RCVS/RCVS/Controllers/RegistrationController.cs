using System.Collections.Generic;
using System.Web.Mvc;
using System.Web.Security;
using RCVS.Classes;
using RCVS.Helpers;
using RCVS.Models;
using RCVS.WebServiceClasses;

namespace RCVS.Controllers
{
	public class RegistrationController : Controller
	{
		public ActionResult Index()
		{
			ViewData["Message"] = "Registration";
			return View();
		}
	}
}
