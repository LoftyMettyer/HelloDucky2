
using System;
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
			var result = xmlHelper.DeserializeFromXmlToObject<AddOrganisationResult>(response);

			var addParameters2 = new AddPositionParameters()
				{
					AddressNumber = 1,
					ContactNumber = result.ContactNumber,
					OrganisationNumber = result.AddressNumber,
					Position = VetName,
					PositionSeniority = (CurrentOrPlanned == CurrentOrPlanned.Current? "C": "P"),
					ValidFrom = DateTime.Now, // (DateTime) StartDate, validation errors!
					ValidTo = System.DateTime.Now // (DateTime) EndDate, validation errors!
				};
			serializedParameters = xmlHelper.SerializeToXml(addParameters2);
			response = client.AddPosition(serializedParameters);

			var Result2 = xmlHelper.DeserializeFromXmlToObject<AddPositionResult>(response);






		}
	}
}