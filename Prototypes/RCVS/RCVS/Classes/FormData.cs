using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Xml.Linq;
using RCVS.Helpers;
using RCVS.IRISWebServices;
using RCVS.WebServiceClasses;

namespace RCVS.Classes
{
	public class FormData
	{
		public enum Forms
		{
			DeclarationOfIntention,
			RenewalOfDeclaration,
			SeeingPractice,
			ExaminationApplicationAndFee
		}

		private readonly List<string> DeclarationOfIntentionActivityIDs = new List<string>
			{
				"0PTD",
				"0TDS",
				"0TPD",
				"0NLC",
				"0UN",
				"0UCC",
				"0UC",
				"0YPE",
				"0PSP",
				"0CSP"
			};

		private readonly List<string> RenewalOfDeclarationActivityIDs = new List<string>
			{
			};

		private readonly List<string> SeeingPracticeActivityIDs = new List<string>
			{
			};

		private readonly List<string> ExaminationApplicationAndFeeActivityIDs = new List<string>
			{
			};

		private Forms _Form;
		private long _ContactNumber;

		public FormData(Forms FormName, long ContactNumber)
		{
			this._Form = FormName;
			this._ContactNumber = ContactNumber;
		}

		public List<SelectContactData_CategoriesResult> GetFormActivities()
		{
			var client = new IRISWebServices.NDataAccessSoapClient(); //Client to call the web services

			var xmlHelper = new XMLHelper(); //XML helper to serialize and deserialize objects
			var selectContactData = new SelectContactDataParameters() { ContactNumber = _ContactNumber };
			var serializedParameters = xmlHelper.SerializeToXml(selectContactData); //Serialize to XML to pass to the web services

			var contactDataSelectionTypes = new IRISWebServices.XMLContactDataSelectionTypes();
			contactDataSelectionTypes = XMLContactDataSelectionTypes.xcdtContactCategories; // Activities

			string response = client.SelectContactData(contactDataSelectionTypes, serializedParameters);

			//Save all the activities for this user (the web services return the whole lot plus its history);
			//further down we need to filter by form and also get only the first (that is, last) instance of each activity
			List<SelectContactData_CategoriesResult> allActivities = new List<SelectContactData_CategoriesResult>();

			SelectContactData_CategoriesResult activity;

			foreach (XElement x in XDocument.Parse(response).Descendants("DataRow"))
			{
				//Dates: always a headache

				DateTime activityDate;
				DateTime.TryParse(x.Element("ActivityDate").Value, out activityDate);

				DateTime validFrom;
				DateTime.TryParse(x.Element("ValidFrom").Value, out validFrom);

				DateTime validTo;
				DateTime.TryParse(x.Element("ValidTo").Value, out validTo);

				DateTime amendedOn;
				DateTime.TryParse(x.Element("AmendedOn").Value, out amendedOn);

				int quantity;
				Int32.TryParse(x.Element("Quantity").Value, out quantity);

				activity = new SelectContactData_CategoriesResult
				 {
					 ContactNumber = Convert.ToInt64(x.Element("ContactNumber").Value),
					 ActivityCode = x.Element("ActivityCode").Value,
					 ActivityValueCode = x.Element("ActivityValueCode").Value,
					 Quantity = quantity,
					 ActivityDate = activityDate,
					 SourceCode = x.Element("SourceCode").Value,
					 ValidFrom = validFrom,
					 ValidTo = validTo,
					 AmendedBy = x.Element("AmendedBy").Value,
					 AmendedOn = amendedOn,
					 Notes = x.Element("Notes").Value,
					 ActivityDesc = x.Element("ActivityDesc").Value,
					 ActivityValueDesc = x.Element("ActivityValueDesc").Value,
					 SourceDesc = x.Element("SourceDesc").Value,
					 RgbActivityValue = x.Element("RgbActivityValue").Value,
					 NoteFlag = x.Element("NoteFlag").Value,
					 Status = x.Element("Status").Value,
					 Access = x.Element("Access").Value,
					 StatusOrder = x.Element("StatusOrder").Value
				 };
				allActivities.Add(activity);
			}

			//We need to initialize this variable even though we will assign a value to it in the switch below;
			//this needs to be done so Visual Studio won't refuse to compile with the message "Use of unassigned local variable 'ActivityIDsToIterateOver'"
			List<string> ActivityIDsToIterateOver = new List<string>();

			switch (_Form)
			{
				case Forms.DeclarationOfIntention:
					ActivityIDsToIterateOver = DeclarationOfIntentionActivityIDs;
					break;
				case Forms.RenewalOfDeclaration:
					ActivityIDsToIterateOver = RenewalOfDeclarationActivityIDs;
					break;
				case Forms.SeeingPractice:
					ActivityIDsToIterateOver = SeeingPracticeActivityIDs;
					break;
				case Forms.ExaminationApplicationAndFee:
					ActivityIDsToIterateOver = ExaminationApplicationAndFeeActivityIDs;
					break;
			}

			//Populate a new filtered list with the activities belonging to the requested form
			List<SelectContactData_CategoriesResult> filteredActivities = new List<SelectContactData_CategoriesResult>();

			foreach (string activityCode in ActivityIDsToIterateOver)
			{
				//If the activity exists, add it to the list
				if (allActivities.FindIndex(a => a.ActivityCode == activityCode) >= 0)
				{
					activity = allActivities.Last(a => a.ActivityCode == activityCode);
					filteredActivities.Add(activity);
				}
			}

			return filteredActivities;
		}
	}
}