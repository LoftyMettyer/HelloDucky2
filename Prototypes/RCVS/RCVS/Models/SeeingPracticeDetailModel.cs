
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
	public class SeeingPracticeDetailModel : PracticeArrangement, iModel
	{
		public long UserID { get; set; }

		public SeeingPracticeDetailModel()
		{
			Address = new Address();
		}


		public void Load()
		{
			Address = new Address();

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

			//Days = Utils.ListOfDays();
			//Months = Utils.ListOfMonths();
			//Years = Utils.ListOfYears();
			Address.Countries = countries;


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