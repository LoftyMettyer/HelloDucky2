using System;
using System.Collections.Generic;
using System.Globalization;
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

		public ActionResult DeclarationOfIntention()
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


			//Contact Number is 571 for nick/nick

			var years = new List<SelectListItem>();
			years.Add(new SelectListItem { Value = "14", Text = "2014" });
			years.Add(new SelectListItem { Value = "15", Text = "2015" });
			years.Add(new SelectListItem { Value = "16", Text = "2016" });

			var model = new DeclarationOfIntentionModel
				{
					YearsDropdown = years,
					Activities = activities
				};

			model.Load();

			return View(model);
		}

			[HttpPost]	
			[ValidateAntiForgeryToken]
		public ActionResult DeclarationOfIntention(DeclarationOfIntentionModel model, FormCollection values)
			{
				//save the Year to Sit
				string response;
				var client = new IRISWebServices.NDataAccessSoapClient(); //Client to call the web services

				var yearToSit = model.YearsDropdown.ToString();

				var XmlHelper = new XMLHelper(); //XML helper to serialize and deserialize objects
				var addActivityParameters = new AddActivityParameters { ContactNumber = 571, Activity = "0YPE", ActivityValue = yearToSit, Source = "WEB" };
				var serializedParameters = XmlHelper.SerializeToXml(addActivityParameters); //Serialize to XML to pass to the web services

				response = client.AddActivity(serializedParameters);

				model.Save();

				return RedirectToAction("Index");
			}

		public ActionResult RenewalOfDeclaration()
		{
			var model = new RenewalOfDeclarationModel();
			model.LoadLookups();

			return View(model);
		}

		[HttpPost]
		public ActionResult RenewalOfDeclaration(RenewalOfDeclarationModel model)
		{
			model.Save();

			model.LoadLookups();

			return View(model);
		}

		public ActionResult SeeingPractice()
		{
			//SeeingPracticeModel = Models.SeeingPracticeModel();
			return View();
		}

		public ActionResult StatutoryMembershipExamination()
		{
			//SeeingPracticeModel = Models.SeeingPracticeModel();
			return View();
		}

		
		//[HttpPost]
		//public ActionResult RenewalOfDeclaration1B()
		//{
		//	return View();
		//}


	}
}
