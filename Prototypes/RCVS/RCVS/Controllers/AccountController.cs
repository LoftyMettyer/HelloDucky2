using System;
using System.Collections.Generic;
using System.Linq;
using System.Transactions;
using System.Web.Mvc;
using System.Web.Security;
using DotNetOpenAuth.AspNet;
using Microsoft.Web.WebPages.OAuth;
using RCVS.Classes;
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

			var loginRegisteredUserParameters = new LoginRegisteredUserParameters { UserName = model.UserName, Password = model.Password }; //Create an object with the login credentials
			var serializedParameters = XmlHelper.SerializeToXml(loginRegisteredUserParameters); //Serialize to XML to pass to the web services

			response = client.LoginRegisteredUser(serializedParameters); //Call the login method

			//If the response message contains "ErrorMessage", deserialize into an ErrorResult object
			if (response.Contains("ErrorMessage"))
			{
				ErrorResult errorResult = XmlHelper.DeserializeFromXmlToObject<ErrorResult>(response);
				// If we got this far, something failed, redisplay form
				ModelState.AddModelError("", errorResult.ErrorMessage);
				return View(model);
			}

			//Deserialize into a LoginRegisteredUserResult object
			LoginRegisteredUserResult loginRegisteredUserResult = XmlHelper.DeserializeFromXmlToObject<LoginRegisteredUserResult>(response);

			//Get tue User details
			var selectContactDataParameters = new SelectContactDataParameters() { ContactNumber = loginRegisteredUserResult.ContactNumber };
			serializedParameters = XmlHelper.SerializeToXml(selectContactDataParameters); //Serialize to XML to pass to the web services

			response = client.SelectContactData(IRISWebServices.XMLContactDataSelectionTypes.xcdtContactInformation, serializedParameters);

			//We don't need all the fields returned by the web services, so instead of casting the result into an object that would need to have every field),
			//we use LINQ to get only what we need

			var temp  = from u in XDocument.Parse(response).Descendants("DataRow")
									select  new User
									{
										ContactNumber = Convert.ToInt64(u.Element("ContactNumber").Value),
										AddressNumber = Convert.ToInt64(u.Element("AddressNumber").Value),
										ContactName = u.Element("ContactName").Value,
										Title = u.Element("Title").Value,
										Initials = u.Element("Initials").Value,
										Forenames = u.Element("Forenames").Value,
										Surname = u.Element("Surname").Value,
										Honorifics = u.Element("Honorifics").Value,
										Salutation = u.Element("Salutation").Value,
										LabelName = u.Element("LabelName").Value,
									};


			User user = (User) temp.FirstOrDefault();
			//SelectContactData_InformationResult selectContactData_InformationResult = XmlHelper.DeserializeFromXmlToObject<SelectContactData_InformationResult>(response);

			Session["User"] = user; //Save the User details in Session
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

			var model = new RegisterModel();
			model.LoadLookups();

			return View(model);
		}

		[HttpPost]
		[AllowAnonymous]
		[ValidateAntiForgeryToken]
		public ActionResult Register(RegisterModel model, FormCollection values)
		{
			model.LoadLookups();

			var client = new IRISWebServices.NDataAccessSoapClient(); //Client to call the web services
			var XmlHelper = new XMLHelper(); //XML helper to serialize and deserialize objects

			var addContactParameters = new AddContactParameters //Create an object with the contact details
			{
				Title = values["Title"],
				Forenames = values["Forenames"],
				Surname = values["Surnames"],
				DateOfBirth = Convert.ToDateTime(values["DOB"] ),
				Address = values["AddressLine1"] + Environment.NewLine + values["AddressLine2"] + Environment.NewLine + values["AddressLine3"],
				Town = values["City"],
				County = values["County"],
				Postcode = values["Postcode"],
				Country = values["Countries"],
				EmailAddress = values["EmailAddress"],
				//UserName = values["EmailAddress"], //This is not an error, UserName should be set to the email address
				//Password = values["Password"],
				Salutation = values["Title"] + " " + values["Surnames"],
				LabelName = values["Title"] + " " + values["Forenames"] + " " + values["Surnames"],
				Status = "WA", //Always "WA" (Web applicant)
				Source = "WEB" //Always "WEB"
			};

			var serializedParameters = XmlHelper.SerializeToXml(addContactParameters); //Serialize to XML to pass to the web services

			string response = client.AddContact(serializedParameters);

			//If the response message contains "ErrorMessage", deserialize into an ErrorResult object
			if (response.Contains("ErrorMessage"))
			{
				ErrorResult errorResult = XmlHelper.DeserializeFromXmlToObject<ErrorResult>(response);
				// If we got this far, something failed, redisplay form
				ModelState.AddModelError("", errorResult.ErrorMessage);
				return View(model);
			}

			//Deserialize into a AddContactResult object
			AddContactResult addContactResult = XmlHelper.DeserializeFromXmlToObject<AddContactResult>(response);
			//Once the contact is added, we need to call another web service to actually add the contact to the list of registered users
			var addRegisteredUserParameters = new AddRegisteredUserParameters
				{
					ContactNumber = addContactResult.ContactNumber,
					UserName = values["EmailAddress"],
					Password = values["Password"],
					EMailAddress = values["EmailAddress"]
				};

			serializedParameters = XmlHelper.SerializeToXml(addRegisteredUserParameters);
			response = client.AddRegisteredUser(serializedParameters);
			// AddRegisteredUserResult addRegisteredUserResult = XmlHelper.DeserializeFromXmlToObject<AddRegisteredUserResult>(response);

			return RedirectToAction("RegistrationSuccessful", "Account");
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
