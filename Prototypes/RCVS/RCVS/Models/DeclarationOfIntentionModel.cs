using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Web.Mvc;
using RCVS.Classes;
using RCVS.Helpers;
using RCVS.Structures;
using RCVS.WebServiceClasses;

namespace RCVS.Models
{
	public class DeclarationOfIntentionModel : BaseModel
	{
		[Required]
		[DisplayName("Select the year in which you plan to sit the statutory membership examination")]
		//public int YearToSit { get; set; }
		public IEnumerable<SelectListItem> YearsDropdown { get; set; }

		public string Footnote1
		{
			get { return "If you do not sit this examination at the next available session or, if you must re-confirm your intention to sit within a reasonable period of time by completing a renewal of intention form"; }
		}

		public string Activity { get; set; }
		public IEnumerable<SelectListItem> Activities { get; set; }


		[DisplayName("Do you plan to 'see practice'?")]
		public bool? PlanToSeePractice { get; set; }

		[DisplayName("Are you currently seeing practice or have you made arrangements?")]
		public bool? CurrentlySeeingPractice { get; set; }

		[DisplayName("TODO - This need to be a file upload of some decription ")]
		public string IELTS { get; set; }

		[DisplayName("When do you plan to take the test?")]
		public DateTime? TakeTestPlanDate { get; set; }

		[DisplayName("If you have taken a test, give details and send your TRF for verification")]
		public TRFDetails TrfDetails { get; set; }

		[DisplayName("Title of primary veterinary degree and recognised abbreviation if any")]
		public Degree PrimaryVetinaryDegree { get; set; }

		public University UniversityAwarded { get; set; }

		[DisplayName("Date of graduation")]
		public DateTime? GraduationDate { get; set; }

		[DisplayName("When did you start your course?")]
		public DateTime? CourseStartDate { get; set; }

		[DisplayName("When did you complete your course?")]
		public DateTime? CourseEndDate { get; set; }

		[DisplayName("What is the normal length of your course?")]
		public Classes.TimePeriod NormalCourseLength { get; set; }

		[DisplayName("Have you enclosed a transcript?")]
		public bool? HasEnclosedTranscript { get; set; }

		[DisplayName("TODO - This need to be a file upload of some decription ")]
		public string EnclosedTranscript { get; set; }

		public override void Load()
		{
			long ContactNumber = 571;

			if (ContactNumber != null)
			{
				//get data
				//NormalCourseLength = new TimePeriod { Months = 1, Years = 2 };


			}
		}

		public override void Save()
		{

			int ContactNumber = 571;

			string response;
			var client = new IRISWebServices.NDataAccessSoapClient();

			var yearToSit = YearsDropdown.ToString();

			var XmlHelper = new XMLHelper(); //XML helper to serialize and deserialize objects

			var addActivityParameters = new AddActivityParameters
			{
				ContactNumber = UserID,
				Activity = "YYGRAD",
				ActivityValue = "MT",
				ActivityDate = GraduationDate,
				Source = "WEB"
			};

			var serializedParameters = XmlHelper.SerializeToXml(addActivityParameters);
			response = client.AddActivity(serializedParameters);

			if (ContactNumber != null)
			{
				// addactivities
			}
			else
			{
				// update activities
			}
		}
	}
}