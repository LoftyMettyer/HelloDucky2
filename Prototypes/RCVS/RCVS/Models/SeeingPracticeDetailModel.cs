
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web.Mvc;
using System.Xml.Linq;
using RCVS.Classes;
using RCVS.Enums;
using RCVS.Helpers;
using RCVS.Interfaces;
using RCVS.WebServiceClasses;

namespace RCVS.Models
{
	public class SeeingPracticeDetailModel : PracticeArrangement
	{
		public long UserID { get; set; }

		public SeeingPracticeDetailModel()
		{
			Address = new Address();
		}


		public void Load(string rowNumber)
		{
			User user = (User)System.Web.HttpContext.Current.Session["User"];
			long UserID = Convert.ToInt64(user.ContactNumber);

			Address = new Address();

			string response;
			var client = new IRISWebServices.NDataAccessSoapClient(); //Client to call the web services

			//Grab a list of countries from the lookup table
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

			Address.Countries = countries;

			////Are we editing a value?
			//if (rowNumber != "")
			//{
			//	//client = new IRISWebServices.NDataAccessSoapClient(); //Client to call the web services
			//	var xmlHelper = new XMLHelper(); //XML helper to serialize and deserialize objects

			//	//Get Position
			//	var selectContactDataParameters = new SelectContactDataParameters() { ContactNumber = UserID };
			//	var serializedParameters = xmlHelper.SerializeToXml(selectContactDataParameters); //Serialize to XML to pass to the web services

			//	response = client.SelectContactData(IRISWebServices.XMLContactDataSelectionTypes.xcdtContactPositions, serializedParameters);

			//	var doc = XDocument.Parse(response);

			//	//var query = from data in doc.Descendants("DataRow")
			//	//						select new PracticeArrangement
			//	//						{
			//	//							PracticeName = (string)data.Element("ContactName"),
			//	//							CurrentOrPlanned = ((string)data.Element("PositionSeniority") == "P" ? CurrentOrPlanned.Planned : CurrentOrPlanned.Current),
			//	//							StartDate = DateTime.ParseExact((string)data.Element("ValidFrom"), "dd/MM/yyyy", null),
			//	//							EndDate = DateTime.ParseExact((string)data.Element("ValidTo"), "dd/MM/yyyy", null),
			//	//							VetName = (string)data.Element("Position")
			//	//						};

			//	//practiceArrangements = query.ToList();		
			//}

		}

		public void Save()
		{

			User user = (User)System.Web.HttpContext.Current.Session["User"];
			long UserID = Convert.ToInt64(user.ContactNumber);

			//save the Year to Sit
			string response;
			var client = new IRISWebServices.NDataAccessSoapClient();

			var xmlHelper = new XMLHelper();
			var addParameters = new AddOrganisationParameters()
				{
					Source = "WEB",
					Name = PracticeName,
					Address = Address.AddressLine1,
					Town = Address.Town,
					County = Address.County,
					Country = Address.Country,
					Postcode = Address.Postcode
				};
			var serializedParameters = xmlHelper.SerializeToXml(addParameters);
			response = client.AddOrganisation(serializedParameters);
			Utils.LogWebServiceCall("AddOrganisation", serializedParameters, response); //Log the call and response
			var result = xmlHelper.DeserializeFromXmlToObject<AddOrganisationResult>(response);

			var addParameters2 = new AddPositionParameters()
				{
					AddressNumber = result.AddressNumber,
					ContactNumber = user.ContactNumber,
					OrganisationNumber = result.ContactNumber,
					Position = VetName,
					PositionSeniority = (CurrentOrPlanned == CurrentOrPlanned.Current? "C": "P"),
					ValidFrom = (DateTime)StartDate,
					ValidTo = (DateTime) EndDate					
				};
			serializedParameters = xmlHelper.SerializeToXml(addParameters2);
			response = client.AddPosition(serializedParameters);
			Utils.LogWebServiceCall("AddPosition", serializedParameters, response); //Log the call and response

			var Result2 = xmlHelper.DeserializeFromXmlToObject<AddPositionResult>(response);

			client.Close();
		}
	}
}