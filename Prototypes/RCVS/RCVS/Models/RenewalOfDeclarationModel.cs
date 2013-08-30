using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Xml.Linq;
using RCVS.Classes;
using RCVS.Helpers;
using RCVS.Structures;
using RCVS.WebServiceClasses;

namespace RCVS.Models
{
	public class RenewalOfDeclarationModel : BaseModel
	{

		public string RenewalReasonCode { get; set; }

		[Required]
		[DisplayName("Reason for renewal of intention to sit?")]
		public List<RenewalReason> RenewalReasons { get; set; }

		//		public ICollection<RenewalReason> RenewalReasons { get; set; }
		//		public IEnumerable<SelectListItem> RenewalReasons { get; set; }

		public ExamAttempts Attempts { get; set; }

		[DisplayName("Year in which you plan to sit your examination")]
		public int PlannedYearToSit { get; set; }

		[DisplayName("Do you plan to 'see practice'?")]
		public string PlanToSeePractice { get; set; }

		[DisplayName("Are you currently seeing practice or have you made arrangements?")]
		public string CurrentlySeeingPractice { get; set; }

		public string IELTS { get; set; }

		[DisplayName("When do you plan to take the test?")]
		public DateTime? PlannedTestDate { get; set; }

		[DisplayName("If you have taken a test, give details")]
		public TRFDetails PreviousTest { get; set; }

		public void LoadLookups()
		{
			var lookupDataType = new IRISWebServices.XMLLookupDataTypes();
			lookupDataType = IRISWebServices.XMLLookupDataTypes.xldtRenewalChangeReasons;

			var client = new IRISWebServices.NDataAccessSoapClient();

			var response = client.GetLookupData(lookupDataType, "");
			Utils.LogWebServiceCall("GetLookupData", "NONE", response); //Log the call and response

			client.Close();

			var doc = XDocument.Parse(response);
			var list = (from xElement in doc.Root.Elements("DataRow")
									select new RenewalReason
										{
											Automatic = xElement.Element("Automatic").Value,
											Description = xElement.Element("RenewalChangeReasonDesc").Value,
											Reason = xElement.Element("RenewalChangeReason").Value,
											Text = xElement.Element("RenewalChangeReasonDesc").Value,
											Value = xElement.Element("RenewalChangeReason").Value
										}).ToList();

			RenewalReasons = list;

		}

		public override void Load()
		{
			User user = (User)System.Web.HttpContext.Current.Session["User"];

			long contactNumber = Convert.ToInt64(user.ContactNumber);

			if (contactNumber != null)
			{
				//Getting the data for the DeclarationOfIntention form in the RenewalOfDeclaration form; THIS IS NOT AN ERROR, we need a piece of data from that form
				FormData formData = new FormData(FormData.Forms.DeclarationOfIntention, contactNumber);
				List<SelectContactData_CategoriesResult> activityList = formData.GetFormActivities();

				if (Utils.ActivityIndex(activityList, "0YPE") >= 0)
				{
					PlannedYearToSit = Convert.ToInt32(activityList.First(activity => activity.ActivityCode == "0YPE").ActivityValueDesc);
				}
			}
		}

		public override void Save()
		{

		}
	}
}