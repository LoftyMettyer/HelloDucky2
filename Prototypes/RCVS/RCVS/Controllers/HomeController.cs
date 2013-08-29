using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Web.Security;
using System.Xml.Linq;
using RCVS.Classes;
using RCVS.Helpers;
using RCVS.IRISWebServices;
using RCVS.Models;
using RCVS.WebServiceClasses;

namespace RCVS.Controllers
{
	[Authorize]
	public class HomeController : Controller
	{
		public ActionResult Index()
		{
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

		[Authorize]
		public ActionResult DeclarationOfIntention()
		{
			ViewBag.SuccessfulSubmit = false;

			int index;
			string response;
			var client = new IRISWebServices.NDataAccessSoapClient(); //Client to call the web services

			//set the lookup key..
			var lookupDataType = new IRISWebServices.XMLLookupDataTypes();
			lookupDataType = XMLLookupDataTypes.xldtActivities; // Activities

			var XmlHelper = new XMLHelper(); //XML helper to serialize and deserialize objects
			var GetLookupDataParameters = new GetLookupDataParameters { };
			var serializedParameters = XmlHelper.SerializeToXml(GetLookupDataParameters); //Serialize to XML to pass to the web services

			response = client.GetLookupData(lookupDataType, serializedParameters);

			//Load the model first; then we need to see if we have a saved value for any of the fields and pre-select it
			DeclarationOfIntentionModel m = new DeclarationOfIntentionModel();
			var model = m.LoadModel();

			var activities = from activity in XDocument.Parse(response).Descendants("DataRow")
											 select new SelectListItem
												 {
													 Value = activity.Element("Activity").Value,
													 Text = activity.Element("ActivityDesc").Value
												 };
			model.Activities = activities;

			//This doesn't need to be pre-selected manually, the framework does it
			var years = new List<SelectListItem>();
			years.Add(new SelectListItem { Value = "14", Text = "2014" });
			years.Add(new SelectListItem { Value = "15", Text = "2015" });
			years.Add(new SelectListItem { Value = "16", Text = "2016" });
			model.YearsDropdown = years;

			//This one DOES need to be pre-selected manually
			var normalCourseLengthDropdown = new List<SelectListItem>();
			normalCourseLengthDropdown.Add(new SelectListItem { Value = "3", Text = "3 years" });
			normalCourseLengthDropdown.Add(new SelectListItem { Value = "4", Text = "4 years" });
			normalCourseLengthDropdown.Add(new SelectListItem { Value = "5", Text = "5 years" });
			normalCourseLengthDropdown.Add(new SelectListItem { Value = "7", Text = "7 years" });
			index = normalCourseLengthDropdown.FindIndex(x => Convert.ToInt32(x.Value) == model.NormalCourseLength); // ... Get the index of the saved value (if any)...
			if (index >= 0)
			{
				normalCourseLengthDropdown[index].Selected = true; // ... and select it
			};

			model.NormalCourseLengthDropdown = normalCourseLengthDropdown;

			//This one DOES need to be pre-selected manually
			var universityThatAwardedDegreeCountriesDropdown = new List<SelectListItem>();
			universityThatAwardedDegreeCountriesDropdown.Add(new SelectListItem { Value = "F", Text = "France" });
			universityThatAwardedDegreeCountriesDropdown.Add(new SelectListItem { Value = "I", Text = "Ireland" });
			universityThatAwardedDegreeCountriesDropdown.Add(new SelectListItem { Value = "IT", Text = "Italy" });
			universityThatAwardedDegreeCountriesDropdown.Add(new SelectListItem { Value = "NL", Text = "Netherlands" });
			universityThatAwardedDegreeCountriesDropdown.Add(new SelectListItem { Value = "SP", Text = "Spain" });
			index = universityThatAwardedDegreeCountriesDropdown.FindIndex(x => x.Value == model.UniversityThatAwardedDegree.Country); // ... Get the index of the saved value (if any)...
			if (index >= 0)
			{
				universityThatAwardedDegreeCountriesDropdown[index].Selected = true; // ... and select it
			}
			;

			var universityThatAwardedDegree = new University
				{
					CountriesDropdown = universityThatAwardedDegreeCountriesDropdown,
					Name = model.UniversityThatAwardedDegree.Name,
					City = model.UniversityThatAwardedDegree.City
				};
			model.UniversityThatAwardedDegree = universityThatAwardedDegree;

			return View(model);
		}

		[HttpPost]
		[ValidateAntiForgeryToken]
		public ActionResult DeclarationOfIntention(DeclarationOfIntentionModel model, FormCollection values)
		{
			model.Save();

			ViewBag.SuccessfulSubmit = true;
			return RedirectToAction("DeclarationOfIntention");
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


		[HttpGet]
		public ActionResult SeeingPractice()
		{
			var model = new SeeingPracticeModel();
			model.Load();

			return View(@model);
		}

		[HttpGet]
		public ActionResult ExaminationApplicationAndFee()
		{
			var model = new ExaminationApplicationAndFeeModel();
			model.Load();

			var yearOfLastApplication = new List<SelectListItem>();
			yearOfLastApplication.Add(new SelectListItem { Value = "N09", Text = "2009" });
			yearOfLastApplication.Add(new SelectListItem { Value = "N10", Text = "2010" });
			yearOfLastApplication.Add(new SelectListItem { Value = "N11", Text = "2011" });
			yearOfLastApplication.Add(new SelectListItem { Value = "N12", Text = "2012" });
			model.YearOfLastApplicationDropDown = yearOfLastApplication;

			var subjects = new List<SelectListItem>();
			subjects.Add(new SelectListItem { Value = "1",Text = "The Horse" });
			subjects.Add(new SelectListItem { Value = "2", Text = "Small Companion Animals" });
			subjects.Add(new SelectListItem { Value = "3", Text = "Production Animals" });
			subjects.Add(new SelectListItem { Value = "4", Text = "Veterinary Public Health" });
			model.SubjectsWithPermissionDropDown = subjects;

			var amountYouArePaying = new List<SelectListItem>();
			amountYouArePaying.Add(new SelectListItem {Value = "1430", Text = "£1430"});
			amountYouArePaying.Add(new SelectListItem { Value = "715", Text = "£715" });
			model.AmountYouArePayingDropDown = amountYouArePaying;

			return View(model);
		}

		[HttpPost]
		[ValidateAntiForgeryToken]
		public ActionResult ExaminationApplicationAndFee(ExaminationApplicationAndFeeModel model)
		{
			model.Save();
			return View(model);
		}


		[HttpGet]
		public ActionResult Activity_Qualification()
		{
			var model = new QualificationModel();
			return View(model);
		}

		[HttpPost]
		[ValidateAntiForgeryToken]
		public ActionResult Activity_Qualification(QualificationModel model, FormCollection values)
		{
			model.Save();
			return View("ExaminationApplicationAndFee");
		}


		[HttpGet]
		public ActionResult SeeingPracticeDetail()
		{
			var model = new SeeingPracticeDetailModel();
			return PartialView(model);
		}

		[HttpGet]
		public ActionResult SeePracticeDetail()
		{
			var model = new SeeingPracticeDetailModel();
			return PartialView(model);
		}

		[HttpPost]
		[ValidateAntiForgeryToken]
		public ActionResult SeeingPracticeDetail(SeeingPracticeDetailModel model, FormCollection values)
		{
			model.Save();
			//SeeingPractice();

			//return Redirect("SeeingPractice");
			return RedirectToAction("DeclarationOfIntention");

			//var model = new SeeingPracticeModel();
			//model.Load();

			//return View("SeeingPractice");
		}

	}
}
