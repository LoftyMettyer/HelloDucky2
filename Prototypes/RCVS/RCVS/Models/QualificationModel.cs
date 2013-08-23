using RCVS.Classes;
using RCVS.Helpers;
using RCVS.Interfaces;
using RCVS.WebServiceClasses;



namespace RCVS.Models
{
	public class QualificationModel :  Qualification, iModel
	{
		public long UserID { get; set; }

		public void Load()
		{
			throw new System.NotImplementedException();
		}

		public void Save()
		{

			string response;
			var client = new IRISWebServices.NDataAccessSoapClient(); 

			var XmlHelper = new XMLHelper();
			var addActivityParameters = new AddActivityParameters { ContactNumber = 571, Activity = "0PSP", ActivityValue = "Y", Source = "WEB", Notes = AwardingBody };
			var serializedParameters = XmlHelper.SerializeToXml(addActivityParameters); 

			response = client.AddActivity(serializedParameters);

		}
	}
}