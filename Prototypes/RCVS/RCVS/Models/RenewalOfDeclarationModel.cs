using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Xml.Linq;
using RCVS.Classes;
using RCVS.Structures;

namespace RCVS.Models
{
	public class RenewalOfDeclarationModel
	{

		public string RenewalReasonCode { get; set; }

		[Required]
		[DisplayName("Reason for renewal of intention to sit?")]
		public List<RenewalReason> RenewalReasons { get; set; }

//		public ICollection<RenewalReason> RenewalReasons { get; set; }
//		public IEnumerable<SelectListItem> RenewalReasons { get; set; }

		public ExamAttempts Attempts { get; set; }

		[DisplayName("Year in which you plan to sit the examination")]
		public int YearToSit { get; set; }

		[DisplayName("Do you plan to 'see practice'?")]
		public bool PlanToSeePractice { get; set; }

		[DisplayName("Are you currently seeing practice or have you made arrangements?")]
		public bool CurrentlySeeingPractice  { get; set; }

		public string IELTS { get; set; }

		[DisplayName("When do you plan to take the test?")]
		public DateTime? PlannedTestDate { get; set; }

		[DisplayName("If you have taken a test, give details")]
		public TRFDetails PreviousTest { get; set; }


		public void Save()
		{
			int _Save = 1;
		}

		public void LoadLookups()
		{
			var lookupDataType = new IRISWebServices.XMLLookupDataTypes();
			lookupDataType = IRISWebServices.XMLLookupDataTypes.xldtRenewalChangeReasons;

			var client = new IRISWebServices.NDataAccessSoapClient();

			var response = client.GetLookupData(lookupDataType, "");

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



	}
}