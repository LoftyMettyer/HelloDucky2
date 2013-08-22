using System;
using System.Collections.Generic;
using System.Linq;
using System.Transactions;
using System.Web.Mvc;
using System.Web.Security;
using DotNetOpenAuth.AspNet;
using Microsoft.Web.WebPages.OAuth;
using RCVS.Helpers;
using RCVS.WebServiceClasses;
using WebMatrix.WebData;
using RCVS.Filters;
using RCVS.Models;
using System.Xml.Linq;

namespace RCVS.Controllers
{
	[Authorize]
	public class AccountController : Controller
	{
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

			//Deserialize into a LoginResult object
			LoginResult loginResult = XmlHelper.DeserializeFromXmlToObject<LoginResult>(response);
			FormsAuthentication.SetAuthCookie(model.UserName, true);

			if (String.IsNullOrEmpty(returnUrl))
			{
				return RedirectToAction("Index", "Home");
			}

			return Redirect(returnUrl);

		}

		[AllowAnonymous]
		public ActionResult Register()
		{
			int i;

			ViewData["Message"] = "Register";

			string response;
			var client = new IRISWebServices.NDataAccessSoapClient(); //Client to call the web services

			//Set the lookup key..
			var lookupDataType = new IRISWebServices.XMLLookupDataTypes();
			lookupDataType = IRISWebServices.XMLLookupDataTypes.xldtCountries; //Countries

			response = client.GetLookupData(lookupDataType, "");

			var countries = new List<SelectListItem>();
			countries.Add(new SelectListItem { Value = "", Text = "" }); //Empty option

			countries.AddRange(from country in XDocument.Parse(response).Descendants("DataRow")
												 select new SelectListItem
												 {
													 Value = country.Element("Country").Value,
													 Text = country.Element("CountryDesc").Value
												 });

			var model = new RegisterModel
			{
				Days = Utils.DropdownList(Utils.DropdownListType.Days),
				Months = Utils.DropdownList(Utils.DropdownListType.Months),
				Years = Utils.DropdownList(Utils.DropdownListType.Years),
				Countries = countries
			};

			return View(model);
		}

		[HttpPost]
		[AllowAnonymous]
		[ValidateAntiForgeryToken]
		public ActionResult Register(RegisterModel model, FormCollection values)
		{

			string response;
			var client = new IRISWebServices.NDataAccessSoapClient(); //Client to call the web services
			var XmlHelper = new XMLHelper(); //XML helper to serialize and deserialize objects

			var addContactParameters = new AddContactParameters //Create an object with the contact details
			{
				Title = values["Title"],
				Forenames = values["Forenames"],
				Surname = values["Surnames"],
				DateOfBirth = Convert.ToDateTime((values["Days"] + "/" + values["Months"] + "/" + values["Years"])),
				Address = values["AddressLine1"] + Environment.NewLine + values["AddressLine2"] + Environment.NewLine + values["AddressLine3"],
				Town = values["City"],
				County = values["County"],
				Postcode = values["Postcode"],
				Country = values["Countries"],
				EmailAddress = values["EmailAddress"],
				UserName = values["EmailAddress"], //This is not an error, UserName should be set to the email address
				Password = values["Password"],
				Salutation = values["Title"] + " " + values["Surnames"],
				LabelName = values["Title"] + " " + values["Forenames"] + " " + values["Surnames"],
				Status = "WA", //Always "WA" (Web applicant)
				Source = "WEB" //Always "WEB"
			};

			var serializedParameters = XmlHelper.SerializeToXml(addContactParameters); //Serialize to XML to pass to the web services

			response = client.AddContact(serializedParameters);

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
				// AddContactResult addContactResult = XmlHelper.DeserializeFromXmlToObject<AddContactResult>(response);
				return RedirectToAction("RegistrationSuccessful", "Account");
			}

			return View();
		}

		[AllowAnonymous]
		public ActionResult RegistrationSuccessful()
		{
			ViewData["Message"] = "Registration Successful";

			return View();
		}

		//
		// POST: /Account/LogOff
		[HttpPost]
		[ValidateAntiForgeryToken]
		public ActionResult LogOff()
		{
			FormsAuthentication.SignOut();

			return RedirectToAction("Index", "Home");
		}

		#region Helpers
		private ActionResult RedirectToLocal(string returnUrl)
		{
			if (Url.IsLocalUrl(returnUrl))
			{
				return Redirect(returnUrl);
			}
			else
			{
				return RedirectToAction("Index", "Home");
			}
		}

		public enum ManageMessageId
		{
			ChangePasswordSuccess,
			SetPasswordSuccess,
			RemoveLoginSuccess,
		}

		internal class ExternalLoginResult : ActionResult
		{
			public ExternalLoginResult(string provider, string returnUrl)
			{
				Provider = provider;
				ReturnUrl = returnUrl;
			}

			public string Provider { get; private set; }
			public string ReturnUrl { get; private set; }

			public override void ExecuteResult(ControllerContext context)
			{
				OAuthWebSecurity.RequestAuthentication(Provider, ReturnUrl);
			}
		}

		private static string ErrorCodeToString(MembershipCreateStatus createStatus)
		{
			// See http://go.microsoft.com/fwlink/?LinkID=177550 for
			// a full list of status codes.
			switch (createStatus)
			{
				case MembershipCreateStatus.DuplicateUserName:
					return "User name already exists. Please enter a different user name.";

				case MembershipCreateStatus.DuplicateEmail:
					return "A user name for that e-mail address already exists. Please enter a different e-mail address.";

				case MembershipCreateStatus.InvalidPassword:
					return "The password provided is invalid. Please enter a valid password value.";

				case MembershipCreateStatus.InvalidEmail:
					return "The e-mail address provided is invalid. Please check the value and try again.";

				case MembershipCreateStatus.InvalidAnswer:
					return "The password retrieval answer provided is invalid. Please check the value and try again.";

				case MembershipCreateStatus.InvalidQuestion:
					return "The password retrieval question provided is invalid. Please check the value and try again.";

				case MembershipCreateStatus.InvalidUserName:
					return "The user name provided is invalid. Please check the value and try again.";

				case MembershipCreateStatus.ProviderError:
					return "The authentication provider returned an error. Please verify your entry and try again. If the problem persists, please contact your system administrator.";

				case MembershipCreateStatus.UserRejected:
					return "The user creation request has been canceled. Please verify your entry and try again. If the problem persists, please contact your system administrator.";

				default:
					return "An unknown error occurred. Please verify your entry and try again. If the problem persists, please contact your system administrator.";
			}
		}
		#endregion

	}
}
