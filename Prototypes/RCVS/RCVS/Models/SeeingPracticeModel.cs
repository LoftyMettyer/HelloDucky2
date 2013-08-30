using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Xml.Linq;
using RCVS.Classes;
using RCVS.Enums;
using RCVS.Helpers;
using RCVS.IRISWebServices;
using RCVS.WebServiceClasses;

namespace RCVS.Models
{
	public class SeeingPracticeModel : BaseModel
	{

	public int PlannedYearToSit { get; set; }

	public List<PracticeArrangement> Practices { get; set; }


		public override void Load()
		{
			var user = (User)System.Web.HttpContext.Current.Session["User"];
			long contactNumber = Convert.ToInt64(user.ContactNumber);

			//See Practice list...

			NDataAccessSoapClient client;
			client = new IRISWebServices.NDataAccessSoapClient(); //Client to call the web services
			var xmlHelper = new XMLHelper(); //XML helper to serialize and deserialize objects

			//Get Position
			var selectContactDataParameters = new SelectContactDataParameters() { ContactNumber = contactNumber };
			var serializedParameters = xmlHelper.SerializeToXml(selectContactDataParameters); //Serialize to XML to pass to the web services

			string response;
			response = client.SelectContactData(IRISWebServices.XMLContactDataSelectionTypes.xcdtContactPositions, serializedParameters);

			var doc = XDocument.Parse(response);

			var query = from data in doc.Descendants("DataRow")
									select new PracticeArrangement
				{
										PracticeName = (string)data.Element("ContactName"),
										CurrentOrPlanned = ((string)data.Element("PositionSeniority") == "P" ? CurrentOrPlanned.Planned : CurrentOrPlanned.Current),
										StartDate = DateTime.ParseExact((string)data.Element("ValidFrom"), "dd/MM/yyyy", null),
										EndDate = DateTime.ParseExact((string)data.Element("ValidTo"), "dd/MM/yyyy", null),
										VetName = (string)data.Element("Position")
									};

			Practices = query.ToList();				


				//TODO retreuve this data from the webservices. The exacty structure of how to do this is a mystery!!!!


				// Retrieve from web service
			client = new IRISWebServices.NDataAccessSoapClient();

				var XmlHelper = new XMLHelper();
				//var addActivityParameters = new FindActions() { UserID = "571", myActions = "0PSP" };
				//var serializedParameters = XmlHelper.SerializeToXml(addActivityParameters);

				//response = client.FindActions(serializedParameters);
			//Utils.LogWebServiceCall("FindActions", serializedParameters, response); //Log the call and response
				//AddActivity(serializedParameters);

			var addParameters = new FindOrganisationsParameters() { UserID = "571", Source = "Web" }; //, Status = "0PSP"};
			var serializedParameters = XmlHelper.SerializeToXml(addParameters);

			//	var lookupDataType = IRISWebServices.XMLLookupDataTypes.xldtActivitiesAndValues;
				response = client.FindOrganisations(serializedParameters);
			Utils.LogWebServiceCall("FindOrganisations", serializedParameters, response); //Log the call and response

			client.Close();

				//var Result = XmlHelper.DeserializeFromXmlToObject<AddOrganisationResult>(response);

			//AddActivity(serializedParameters);
		}

		public override void Save()
		{

		}
	}
}