using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Web.Mvc;
using System.Web.Security;
using RCVS.Classes;
using RCVS.Helpers;
using RCVS.WebServiceClasses;
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
			Utils.LogWebServiceCall("LoginRegisteredUser", serializedParameters, response); //Log the call and response

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

			//Get the User details
			var selectContactDataParameters = new SelectContactDataParameters() { ContactNumber = loginRegisteredUserResult.ContactNumber };
			serializedParameters = XmlHelper.SerializeToXml(selectContactDataParameters); //Serialize to XML to pass to the web services

			response = client.SelectContactData(IRISWebServices.XMLContactDataSelectionTypes.xcdtContactInformation, serializedParameters);
			Utils.LogWebServiceCall("SelectContactData", serializedParameters, response); //Log the call and response

			//We don't need all the fields returned by the web services, so instead of casting the result into an object that would need to have every field),
			//we use LINQ to get only what we need

			var temp = from u in XDocument.Parse(response).Descendants("DataRow")
								 select new User
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

			User user = (User)temp.FirstOrDefault();
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
				DateOfBirth = Convert.ToDateTime(values["DOB"]),
				Address = values["AddressLine1"] + Environment.NewLine + values["AddressLine2"] + Environment.NewLine + values["AddressLine3"],
				Town = values["City"],
				County = values["County"],
				Postcode = values["Postcode"],
				Country = values["Countries"],
				EmailAddress = values["EmailAddress"],
				Salutation = values["Title"] + " " + values["Surnames"],
				LabelName = values["Title"] + " " + values["Forenames"] + " " + values["Surnames"],
				Status = "WA", //Always "WA" (Web applicant)
				Source = "WEB" //Always "WEB"
			};

			var serializedParameters = XmlHelper.SerializeToXml(addContactParameters); //Serialize to XML to pass to the web services

			string response = client.AddContact(serializedParameters);
			Utils.LogWebServiceCall("AddContact", serializedParameters, response); //Log the call and response

			//If the response message contains "ErrorMessage", deserialize into an ErrorResult object
			if (response.Contains("ErrorMessage"))
			{
				ErrorResult errorResult = XmlHelper.DeserializeFromXmlToObject<ErrorResult>(response);
				// If we got this far, something failed, redisplay form
				ModelState.AddModelError("", errorResult.ErrorMessage);
				client.Close();
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
			Utils.LogWebServiceCall("AddRegisteredUser", serializedParameters, response); //Log the call and response
			// AddRegisteredUserResult addRegisteredUserResult = XmlHelper.DeserializeFromXmlToObject<AddRegisteredUserResult>(response);

			//Trigger an action for this user
			var addActionFromTemplateParameters = new AddActionFromTemplateParameters
				{
					ActionNumber = 638,
					ContactNumber = addContactResult.ContactNumber
				};
			serializedParameters = XmlHelper.SerializeToXml(addActionFromTemplateParameters); //Serialize to XML to pass to the web services
			response = client.AddActionFromTemplate(serializedParameters);
			Utils.LogWebServiceCall("AddActionFromTemplate", serializedParameters, response); //Log the call and response

			client.Close();

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
			System.Web.HttpContext.Current.Session["User"] = null;

			return RedirectToAction("Index", "Home");
		}

		[AllowAnonymous]
		public ActionResult AddressLookup()
		{
			return View();
		}

		[AllowAnonymous]
		public string QASAddressLookup(string postcode)
		{
			//QASClient.QASClient client = new QASClient.QASClient(System.Configuration.ConfigurationManager.AppSettings["com.qas.proweb.serverURL"]);
			//List<string> list = client.SearchByPostcode(postcode);

			var list = new List<string>();

			switch (postcode.ToUpper())
			{
				case "W3 6PF":
					list.Add("57a Grafton Road,London,W3 6PF");
					list.Add("64 Grafton Road,London,W3 6PF");
					list.Add("65 Grafton Road,London,W3 6PF");
					list.Add("66 Grafton Road,London,W3 6PF");
					list.Add("Flat A-b 67 Grafton Road,London,W3 6PF");
					list.Add("Flat C 67 Grafton Road,London,W3 6PF");
					list.Add("68 Grafton Road,London,W3 6PF");
					list.Add("68e Grafton Road,London,W3 6PF");
					list.Add("69 Grafton Road,London,W3 6PF");
					list.Add("70 Grafton Road,London,W3 6PF");
					list.Add("71 Grafton Road,London,W3 6PF");
					list.Add("72 Grafton Road,London,W3 6PF");
					list.Add("73 Grafton Road,London,W3 6PF");
					list.Add("74 Grafton Road,London,W3 6PF");
					list.Add("75 Grafton Road,London,W3 6PF");
					list.Add("75a Grafton Road,London,W3 6PF");
					list.Add("76 Grafton Road,London,W3 6PF");
					list.Add("77 Grafton Road,London,W3 6PF");
					list.Add("78 Grafton Road,London,W3 6PF");
					list.Add("79 Grafton Road,London,W3 6PF");
					list.Add("80 Grafton Road,London,W3 6PF");
					list.Add("80a Grafton Road,London,W3 6PF");
					list.Add("81 Grafton Road,London,W3 6PF");
					break;
				case "EH6 8PF":
					list.Add("3/1 Dalmeny Street,Edinburgh,EH6 8PF");
					list.Add("3/2 Dalmeny Street,Edinburgh,EH6 8PF");
					list.Add("3/3 Dalmeny Street,Edinburgh,EH6 8PF");
					list.Add("3/4 Dalmeny Street,Edinburgh,EH6 8PF");
					list.Add("3/5 Dalmeny Street,Edinburgh,EH6 8PF");
					list.Add("3/6 Dalmeny Street,Edinburgh,EH6 8PF");
					list.Add("3/7 Dalmeny Street,Edinburgh,EH6 8PF");
					list.Add("3/8 Dalmeny Street,Edinburgh,EH6 8PF");
					list.Add("3/9 Dalmeny Street,Edinburgh,EH6 8PF");
					list.Add("7 Dalmeny Street,Edinburgh,EH6 8PF");
					list.Add("9 Dalmeny Street,Edinburgh,EH6 8PF");
					list.Add("11/1 Dalmeny Street,Edinburgh,EH6 8PF");
					list.Add("11/2 Dalmeny Street,Edinburgh,EH6 8PF");
					list.Add("11/3 Dalmeny Street,Edinburgh,EH6 8PF");
					list.Add("11/4 Dalmeny Street,Edinburgh,EH6 8PF");
					list.Add("11/5 Dalmeny Street,Edinburgh,EH6 8PF");
					list.Add("11/6 Dalmeny Street,Edinburgh,EH6 8PF");
					list.Add("11/7 Dalmeny Street,Edinburgh,EH6 8PF");
					list.Add("11/8 Dalmeny Street,Edinburgh,EH6 8PF");
					list.Add("11/9 Dalmeny Street,Edinburgh,EH6 8PF");
					list.Add("11/10 Dalmeny Street,Edinburgh,EH6 8PF");
					list.Add("13/1 Dalmeny Street,Edinburgh,EH6 8PF");
					list.Add("13/2 Dalmeny Street,Edinburgh,EH6 8PF");
					list.Add("13/3 Dalmeny Street,Edinburgh,EH6 8PF");
					list.Add("13/4 Dalmeny Street,Edinburgh,EH6 8PF");
					list.Add("13/5 Dalmeny Street,Edinburgh,EH6 8PF");
					list.Add("13/6 Dalmeny Street,Edinburgh,EH6 8PF");
					list.Add("13/7 Dalmeny Street,Edinburgh,EH6 8PF");
					list.Add("13/8 Dalmeny Street,Edinburgh,EH6 8PF");
					list.Add("13/9 Dalmeny Street,Edinburgh,EH6 8PF");
					break;
				case "OL8 3DF":
					list.Add("137 Hollins Road,Oldham,OL8 3DF");
					list.Add("139 Hollins Road,Oldham,OL8 3DF");
					list.Add("141 Hollins Road,Oldham,OL8 3DF");
					list.Add("143 Hollins Road,Oldham,OL8 3DF");
					list.Add("145 Hollins Road,Oldham,OL8 3DF");
					list.Add("147 Hollins Road,Oldham,OL8 3DF");
					list.Add("149 Hollins Road,Oldham,OL8 3DF");
					list.Add("151 Hollins Road,Oldham,OL8 3DF");
					list.Add("153 Hollins Road,Oldham,OL8 3DF");
					list.Add("155 Hollins Road,Oldham,OL8 3DF");
					list.Add("157 Hollins Road,Oldham,OL8 3DF");
					list.Add("157a Hollins Road,Oldham,OL8 3DF");
					list.Add("159 Hollins Road,Oldham,OL8 3DF");
					list.Add("161 Hollins Road,Oldham,OL8 3DF");
					break;
				case "SE11 4PT":
					list.Add("365 Kennington Road,London,SE11 4PT");
					list.Add("377 Kennington Road,London,SE11 4PT");
					list.Add("379 Kennington Road,London,SE11 4PT");
					list.Add("381 Kennington Road,London,SE11 4PT");
					list.Add("405 Kennington Road,London,SE11 4PT");
					list.Add("412 Kennington Road,London,SE11 4PT");
					list.Add("Flat 414 Kennington Road,London,SE11 4PT");
					break;
			}

			//list.Add("A1,A2,A3,P1");
			//list.Add("B1,B2,B3,P2");
			//list.Add("C1,C2,C3,P3");

			var JSONSerializer = new System.Web.Script.Serialization.JavaScriptSerializer();

			string retVal = JSONSerializer.Serialize(list);
			return retVal;
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

		#endregion

	}
}
