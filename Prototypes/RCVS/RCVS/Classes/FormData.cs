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
			StatutoryMemberShipexamination
		}

		public readonly List<string> DeclarationOfIntentionActivityIDs = new List<string>
			{
				"0PTD",
				"0TDS"
			};

		private Forms form;

		public FormData(Forms FormName)
		{
			this.form = FormName;
		}

		public List<SelectContactData_CategoriesResult> GetFormActivities(long ContactNumber)
		{
			var client = new IRISWebServices.NDataAccessSoapClient(); //Client to call the web services

			var xmlHelper = new XMLHelper(); //XML helper to serialize and deserialize objects
			var selectContactData = new SelectContactDataParameters() { ContactNumber = ContactNumber };
			var serializedParameters = xmlHelper.SerializeToXml(selectContactData); //Serialize to XML to pass to the web services

			var contactDataSelectionTypes = new IRISWebServices.XMLContactDataSelectionTypes();
			contactDataSelectionTypes = XMLContactDataSelectionTypes.xcdtContactCategories; // Activities

			string response = client.SelectContactData(contactDataSelectionTypes, serializedParameters);

			//Save all the activities for this user; further down, 
			List<SelectContactData_CategoriesResult> allActivities = new List<SelectContactData_CategoriesResult>();

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

				SelectContactData_CategoriesResult result = new SelectContactData_CategoriesResult
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
				allActivities.Add(result);
			}


			int asdfasdf = 0;


			return allActivities;
		}
	}
}