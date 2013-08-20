using System.Web.Mvc;
using System.Web.Security;
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

		public ActionResult Form1A()
		{
			int i;

			ViewData["Message"] = "Form 1A";

			////Days
			//var _daysList = new List<SelectListItem>();
			//_daysList.Add(new SelectListItem { Text = "", Value = "" });
			//for (i = 1; i <= 31; i++)
			//{
			//	_daysList.Add(new SelectListItem { Text = i.ToString(), Value = i.ToString() });
			//}

			//ViewBag.DaysList = new SelectList(_daysList, "Value", "Text");

			//	private List<Day> _days
			//		{
			//			get
			//			{
			//				int i;
			//				//Days
			//				var _daysList = new List<Day>();
			//				_daysList.Add(new Day() { ID = -1, Value = "" });
			//				for (i = 1; i <= 31; i++)
			//				{
			//					_daysList.Add(new Day() { ID = i, Value = i.ToString() });
			//				}

			//				return _daysList;
			//			}
			//		}

			//		public IEnumerable<SelectListItem> Days
			//		{
			//			get { return new SelectList(_days, "ID", "Value"); }
			//		}
			//internal class Day
			//{
			//	public int ID { get; set; }
			//	public string Value { get; set; }
			//}


			////Months
			//List<SelectListItem> _monthsList = new List<SelectListItem>();
			//_monthsList.Add(new SelectListItem() { Value = "", Text = "" });
			//for (i = 0; i <= 11; i++)
			//{
			//	_monthsList.Add(new SelectListItem() { Value = (i + 1).ToString(), Text = CultureInfo.CurrentUICulture.DateTimeFormat.MonthNames[i] });
			//}

			////Years
			//List<SelectListItem> _yearsList = new List<SelectListItem>();
			//_yearsList.Add(new SelectListItem() { Value = "", Text = "" });
			//for (i = DateTime.Now.Year; i >= 1900; i--)
			//{
			//	_yearsList.Add(new SelectListItem() { Value = i.ToString(), Text = i.ToString() });
			//}

			//Form1AModel m = new Form1AModel
			//	{
			//		DaysList = new SelectList(_daysList, "Value", "Text"),
			//		MonthsList = new SelectList(_monthsList, "Value", "Text"),
			//		YearsList = new SelectList(_yearsList, "Value", "Text")
			//	};

			//return View(Form1AModel);

			/////////////////////////////////////////////////////////

			return View();
		}

		//[HttpPost]
		//public ActionResult Form1A(FormCollection values)
		//{

		//}


		//
		// GET: /Account/Login

		[AllowAnonymous]
		public ActionResult Login(string returnUrl)
		{
			ViewBag.ReturnUrl = returnUrl;

			return View();
		}

		//
		// POST: /Account/Login

		[HttpPost]
		[AllowAnonymous]
		[ValidateAntiForgeryToken]
		public ActionResult Login(LoginModel model, string returnUrl)
		{
			string response;
			var client = new IRISWebServices.NDataAccessSoapClient(); //Client to call the web services
			var XmlHelper = new XMLHelper(); //XML helper to serialize and deserialize objects

			var loginParameters = new LoginParameters { UserName = model.UserName, Password = model.Password }; //Create an object with the login credentials
			var serializedParameters = XmlHelper.SerializeToXml(loginParameters); //Serialize to XML to pass to the web services

			response = client.Login(serializedParameters); //Call the login method

			//If the response message contains "ErrorMessage", deserialize into an ErrorResult object
			if (response.Contains("ErrorMessage"))
			{
				ErrorResult errorResult = XmlHelper.DeserializeFromXmlToObject<ErrorResult>(response);
				// If we got this far, something failed, redisplay form
				ModelState.AddModelError("", errorResult.ErrorMessage);
				return View(model);
			}
			else //Deserialize into a LoginResult object
			{
				LoginResult loginResult = XmlHelper.DeserializeFromXmlToObject<LoginResult>(response);
				FormsAuthentication.SetAuthCookie(model.UserName, true);

				return RedirectToAction("Form1A", "Registration");
			}
		}
	}
}
