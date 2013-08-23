using System.Collections.Generic;
using System.Collections.ObjectModel;
using RCVS.Classes;
using RCVS.Enums;
using RCVS.Helpers;
using RCVS.WebServiceClasses;

namespace RCVS.Models
{
	public class SeeingPracticeModel : BaseModel
	{

	public int PlannedYearToSit { get; set; }

	public List<PracticeArrangement> Practices { get; set; }


		public override void Load()
		{

				Practices = new List<PracticeArrangement>();

				Practices.Add(new PracticeArrangement
				{
					PracticeName = "Dogs R Us",
					CurrentOrPlanned = CurrentOrPlanned.Planned
					,
					StartDate = System.DateTime.Now,
					EndDate = System.DateTime.Now,
					VetName = "Harry Sullivan"
					,
					Address = new Address(){ AddressLine1 = "12 Windermere Road", AddressLine2 = "Tonypandy", Town = "Rhondda", Postcode = "CF11 ABC" }
				});

				Practices.Add(new PracticeArrangement
				{
					PracticeName = "Cats 4 You",
					CurrentOrPlanned = CurrentOrPlanned.Current
					,
					StartDate = System.DateTime.Now,
					EndDate = System.DateTime.Now,
					VetName = "Ian Chesterton"
					,
					Address = new Address() { AddressLine1 = "132 Kendal Drive", AddressLine2 = "", Town = "Hemel Hempstead", Postcode = "HS33 1VC" }
				});


				//TODO retreuve this data from the webservices. The exacty structure of how to do this is a mystery!!!!


				// Retrieve from web service
				string response;
				var client = new IRISWebServices.NDataAccessSoapClient();

				var XmlHelper = new XMLHelper();
				//var addActivityParameters = new FindActions() { UserID = "571", myActions = "0PSP" };
				//var serializedParameters = XmlHelper.SerializeToXml(addActivityParameters);

				//response = client.FindActions(serializedParameters);
				//AddActivity(serializedParameters);

			var addParameters = new FindOrganisationsParameters() {UserID = "571", Source = "Web"}; //, Status = "0PSP"};
				var serializedParameters = XmlHelper.SerializeToXml(addParameters);




			//	var lookupDataType = IRISWebServices.XMLLookupDataTypes.xldtActivitiesAndValues;
				response = client.FindOrganisations(serializedParameters);

				//var Result = XmlHelper.DeserializeFromXmlToObject<AddOrganisationResult>(response);

			//AddActivity(serializedParameters);









		}

		public override void Save()
		{

		}
	}
}