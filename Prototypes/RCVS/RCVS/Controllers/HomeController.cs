using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Web.Security;
using System.Xml.Linq;
using RCVS.Helpers;
using RCVS.IRISWebServices;
using RCVS.Models;
using RCVS.WebServiceClasses;

namespace RCVS.Controllers
{
	public class HomeController : Controller
	{
		public ActionResult Index()
		{
			ViewBag.Message = "Modify this template to jump-start your ASP.NET MVC application.";

			return View();
		}

		public ActionResult About()
		{
			ViewBag.Message = "Your app description page.";

			return View();
		}

		public ActionResult Contact()
		{
			ViewBag.Message = "Your contact page.";

			return View();
		}

		public ActionResult Form1A2()
		{
			string response;
			var client = new IRISWebServices.NDataAccessSoapClient(); //Client to call the web services

			//set the lookup key..
			var lookupDataType = new IRISWebServices.XMLLookupDataTypes();						
			lookupDataType = XMLLookupDataTypes.xldtActivities; // Activities
			
			var XmlHelper = new XMLHelper(); //XML helper to serialize and deserialize objects
			var GetLookupDataParameters = new GetLookupDataParameters {};
			var serializedParameters = XmlHelper.SerializeToXml(GetLookupDataParameters); //Serialize to XML to pass to the web services

			response = client.GetLookupData(lookupDataType, serializedParameters);
			
			var activities = from activity in XDocument.Parse(response).Descendants("DataRow")
			                 select new SelectListItem
				                 {
					                 Value = activity.Element("Activity").Value,
					                 Text = activity.Element("ActivityDesc").Value
				                 };

			var model = new Form1A2Model
				{
					Activities = activities
				};
	
			return View(model);
		}



		public ActionResult SeeingPractice()
		{
			//SeeingPracticeModel = Models.SeeingPracticeModel();
			return View();
		}
	}
}
