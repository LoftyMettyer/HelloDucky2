﻿using System;
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
			else //Deserialize into a LoginResult object
			{
				LoginResult loginResult = XmlHelper.DeserializeFromXmlToObject<LoginResult>(response);
				FormsAuthentication.SetAuthCookie(model.UserName, true);

				if (String.IsNullOrEmpty(returnUrl))
				{
					return RedirectToAction("Index", "Home");
				}
				else
				{
					return Redirect(returnUrl);
				}
			}
		}

		public ActionResult Registration()
		{
			int i;

			ViewData["Message"] = "Registration Form";

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

			var model = new RegistrationModel
			{
				Days = Utils.DropdownList(Utils.DropdownListType.Days),
				Months = Utils.DropdownList(Utils.DropdownListType.Months),
				Years = Utils.DropdownList(Utils.DropdownListType.Years),
				Countries = countries
			};

			return View(model);
		}

		[HttpPost]
		public ActionResult Registration(RegistrationModel model, FormCollection values)
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

		#region "NOT USED"

		//
		//// GET: /Account/Register

		//[AllowAnonymous]
		//public ActionResult Register()
		//{
		//	return View();
		//}

		//
		// POST: /Account/Register

		//[HttpPost]
		//[AllowAnonymous]
		//[ValidateAntiForgeryToken]
		//public ActionResult Register(RegisterModel model)
		//{
		//	if (ModelState.IsValid)
		//	{
		//		// Attempt to register the user
		//		try
		//		{
		//			WebSecurity.CreateUserAndAccount(model.UserName, model.Password);
		//			WebSecurity.Login(model.UserName, model.Password);
		//			return RedirectToAction("Index", "Home");
		//		}
		//		catch (MembershipCreateUserException e)
		//		{
		//			ModelState.AddModelError("", ErrorCodeToString(e.StatusCode));
		//		}
		//	}

		//	// If we got this far, something failed, redisplay form
		//	return View(model);
		//}

		//
		// POST: /Account/Disassociate

		[HttpPost]
		[ValidateAntiForgeryToken]
		public ActionResult Disassociate(string provider, string providerUserId)
		{
			string ownerAccount = OAuthWebSecurity.GetUserName(provider, providerUserId);
			ManageMessageId? message = null;

			// Only disassociate the account if the currently logged in user is the owner
			if (ownerAccount == User.Identity.Name)
			{
				// Use a transaction to prevent the user from deleting their last login credential
				using (var scope = new TransactionScope(TransactionScopeOption.Required, new TransactionOptions { IsolationLevel = IsolationLevel.Serializable }))
				{
					bool hasLocalAccount = OAuthWebSecurity.HasLocalAccount(WebSecurity.GetUserId(User.Identity.Name));
					if (hasLocalAccount || OAuthWebSecurity.GetAccountsFromUserName(User.Identity.Name).Count > 1)
					{
						OAuthWebSecurity.DeleteAccount(provider, providerUserId);
						scope.Complete();
						message = ManageMessageId.RemoveLoginSuccess;
					}
				}
			}

			return RedirectToAction("Manage", new { Message = message });
		}

		//
		// GET: /Account/Manage

		public ActionResult Manage(ManageMessageId? message)
		{
			ViewBag.StatusMessage =
					message == ManageMessageId.ChangePasswordSuccess ? "Your password has been changed."
					: message == ManageMessageId.SetPasswordSuccess ? "Your password has been set."
					: message == ManageMessageId.RemoveLoginSuccess ? "The external login was removed."
					: "";
			ViewBag.HasLocalPassword = OAuthWebSecurity.HasLocalAccount(WebSecurity.GetUserId(User.Identity.Name));
			ViewBag.ReturnUrl = Url.Action("Manage");
			return View();
		}

		//
		// POST: /Account/Manage

		[HttpPost]
		[ValidateAntiForgeryToken]
		public ActionResult Manage(LocalPasswordModel model)
		{
			bool hasLocalAccount = OAuthWebSecurity.HasLocalAccount(WebSecurity.GetUserId(User.Identity.Name));
			ViewBag.HasLocalPassword = hasLocalAccount;
			ViewBag.ReturnUrl = Url.Action("Manage");
			if (hasLocalAccount)
			{
				if (ModelState.IsValid)
				{
					// ChangePassword will throw an exception rather than return false in certain failure scenarios.
					bool changePasswordSucceeded;
					try
					{
						changePasswordSucceeded = WebSecurity.ChangePassword(User.Identity.Name, model.OldPassword, model.NewPassword);
					}
					catch (Exception)
					{
						changePasswordSucceeded = false;
					}

					if (changePasswordSucceeded)
					{
						return RedirectToAction("Manage", new { Message = ManageMessageId.ChangePasswordSuccess });
					}
					else
					{
						ModelState.AddModelError("", "The current password is incorrect or the new password is invalid.");
					}
				}
			}
			else
			{
				// User does not have a local password so remove any validation errors caused by a missing
				// OldPassword field
				ModelState state = ModelState["OldPassword"];
				if (state != null)
				{
					state.Errors.Clear();
				}

				if (ModelState.IsValid)
				{
					try
					{
						WebSecurity.CreateAccount(User.Identity.Name, model.NewPassword);
						return RedirectToAction("Manage", new { Message = ManageMessageId.SetPasswordSuccess });
					}
					catch (Exception)
					{
						ModelState.AddModelError("", String.Format("Unable to create local account. An account with the name \"{0}\" may already exist.", User.Identity.Name));
					}
				}
			}

			// If we got this far, something failed, redisplay form
			return View(model);
		}

		//
		// POST: /Account/ExternalLogin

		[HttpPost]
		[AllowAnonymous]
		[ValidateAntiForgeryToken]
		public ActionResult ExternalLogin(string provider, string returnUrl)
		{
			return new ExternalLoginResult(provider, Url.Action("ExternalLoginCallback", new { ReturnUrl = returnUrl }));
		}

		//
		// GET: /Account/ExternalLoginCallback

		[AllowAnonymous]
		public ActionResult ExternalLoginCallback(string returnUrl)
		{
			AuthenticationResult result = OAuthWebSecurity.VerifyAuthentication(Url.Action("ExternalLoginCallback", new { ReturnUrl = returnUrl }));
			if (!result.IsSuccessful)
			{
				return RedirectToAction("ExternalLoginFailure");
			}

			if (OAuthWebSecurity.Login(result.Provider, result.ProviderUserId, createPersistentCookie: false))
			{
				return RedirectToLocal(returnUrl);
			}

			if (User.Identity.IsAuthenticated)
			{
				// If the current user is logged in add the new account
				OAuthWebSecurity.CreateOrUpdateAccount(result.Provider, result.ProviderUserId, User.Identity.Name);
				return RedirectToLocal(returnUrl);
			}
			else
			{
				// User is new, ask for their desired membership name
				string loginData = OAuthWebSecurity.SerializeProviderUserId(result.Provider, result.ProviderUserId);
				ViewBag.ProviderDisplayName = OAuthWebSecurity.GetOAuthClientData(result.Provider).DisplayName;
				ViewBag.ReturnUrl = returnUrl;
				return View("ExternalLoginConfirmation", new RegisterExternalLoginModel { UserName = result.UserName, ExternalLoginData = loginData });
			}
		}

		//
		// POST: /Account/ExternalLoginConfirmation

		[HttpPost]
		[AllowAnonymous]
		[ValidateAntiForgeryToken]
		public ActionResult ExternalLoginConfirmation(RegisterExternalLoginModel model, string returnUrl)
		{
			string provider = null;
			string providerUserId = null;

			if (User.Identity.IsAuthenticated || !OAuthWebSecurity.TryDeserializeProviderUserId(model.ExternalLoginData, out provider, out providerUserId))
			{
				return RedirectToAction("Manage");
			}

			if (ModelState.IsValid)
			{
				// Insert a new user into the database
				using (UsersContext db = new UsersContext())
				{
					UserProfile user = db.UserProfiles.FirstOrDefault(u => u.UserName.ToLower() == model.UserName.ToLower());
					// Check if user already exists
					if (user == null)
					{
						// Insert name into the profile table
						db.UserProfiles.Add(new UserProfile { UserName = model.UserName });
						db.SaveChanges();

						OAuthWebSecurity.CreateOrUpdateAccount(provider, providerUserId, model.UserName);
						OAuthWebSecurity.Login(provider, providerUserId, createPersistentCookie: false);

						return RedirectToLocal(returnUrl);
					}
					else
					{
						ModelState.AddModelError("UserName", "User name already exists. Please enter a different user name.");
					}
				}
			}

			ViewBag.ProviderDisplayName = OAuthWebSecurity.GetOAuthClientData(provider).DisplayName;
			ViewBag.ReturnUrl = returnUrl;
			return View(model);
		}

		//
		// GET: /Account/ExternalLoginFailure

		[AllowAnonymous]
		public ActionResult ExternalLoginFailure()
		{
			return View();
		}

		[AllowAnonymous]
		[ChildActionOnly]
		public ActionResult ExternalLoginsList(string returnUrl)
		{
			ViewBag.ReturnUrl = returnUrl;
			return PartialView("_ExternalLoginsListPartial", OAuthWebSecurity.RegisteredClientData);
		}

		[ChildActionOnly]
		public ActionResult RemoveExternalLogins()
		{
			ICollection<OAuthAccount> accounts = OAuthWebSecurity.GetAccountsFromUserName(User.Identity.Name);
			List<ExternalLogin> externalLogins = new List<ExternalLogin>();
			foreach (OAuthAccount account in accounts)
			{
				AuthenticationClientData clientData = OAuthWebSecurity.GetOAuthClientData(account.Provider);

				externalLogins.Add(new ExternalLogin
				{
					Provider = account.Provider,
					ProviderDisplayName = clientData.DisplayName,
					ProviderUserId = account.ProviderUserId,
				});
			}

			ViewBag.ShowRemoveButton = externalLogins.Count > 1 || OAuthWebSecurity.HasLocalAccount(WebSecurity.GetUserId(User.Identity.Name));
			return PartialView("_RemoveExternalLoginsPartial", externalLogins);
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
		#endregion
	}
}
