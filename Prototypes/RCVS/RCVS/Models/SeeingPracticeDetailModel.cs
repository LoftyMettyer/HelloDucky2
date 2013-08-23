
using System;
using RCVS.Classes;
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

		}

		public void Save()
		{

			UserID = 571;

			//save the Year to Sit
			string response;
			var client = new IRISWebServices.NDataAccessSoapClient();

			var XmlHelper = new XMLHelper();
			var addParameters = new AddOrganisationParameters()
				{
					Source = "Web",
					Name = "vetname", // PracticeName,
					Address = "adr1", // Address.AddressLine1,
					Town = "tyown", //Address.Town,
					County = "county", // Address.County,
					Country = "UK", //Address.Country,
					Postcode = Address.Postcode
					//Status = "0PSP"
				};
			var serializedParameters = XmlHelper.SerializeToXml(addParameters);
			response = client.AddOrganisation(serializedParameters);
			var Result = XmlHelper.DeserializeFromXmlToObject<AddOrganisationResult>(response);

			var addParameters2 = new AddPositionParameters()
				{
					AddressNumber = 1,
					ContactNumber = Result.ContactNumber,
					OrganisationNumber = Result.AddressNumber,
					Position = VetName,
					PositionSeniority = "C",
					ValidFrom = DateTime.Now, // (DateTime) StartDate, validation errors!
					ValidTo = System.DateTime.Now // (DateTime) EndDate, validation errors!
				};
			serializedParameters = XmlHelper.SerializeToXml(addParameters2);
			response = client.AddPosition(serializedParameters);

			var Result2 = XmlHelper.DeserializeFromXmlToObject<AddPositionResult>(response);






		}
	}
}