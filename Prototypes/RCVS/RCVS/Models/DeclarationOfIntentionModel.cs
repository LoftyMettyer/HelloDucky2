using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using RCVS.Classes;
using RCVS.Helpers;
using RCVS.Structures;
using RCVS.WebServiceClasses;

namespace RCVS.Models
{
	public class DeclarationOfIntentionModel : BaseModel
	{
		public int YearToSit { get; set; } //To holde the value of the selected year to sit

		[Required]
		[DisplayName("Select the year in which you plan to sit the statutory membership examination")]
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

		[DisplayName("Upload your IELTS test report form here")]
		public HttpPostedFileBase IELTS { get; set; }

		[DisplayName("When do you plan to take the test?")]
		public DateTime TakeTestPlanDate { get; set; }

		[DisplayName("If you have taken a test, give details and send your TRF for verification")]
		public TRFDetails TrfDetails { get; set; }

		[DisplayName("Title of primary veterinary degree and recognised abbreviation if any")]
		public Degree PrimaryVetinaryDegree { get; set; }

		public University UniversityAwarded { get; set; }

		[DisplayName("Date of graduation")]
		public DateTime GraduationDate { get; set; }

		[DisplayName("When did you start your course?")]
		public DateTime? CourseStartDate { get; set; }

		[DisplayName("When did you complete your course?")]
		public DateTime? CourseEndDate { get; set; }

		[DisplayName("What is the normal length of your course?")]
		public TimePeriod NormalCourseLength { get; set; }

		[DisplayName("Have you enclosed a transcript?")]
		public bool? HasEnclosedTranscript { get; set; }

		[DisplayName("TODO - This need to be a file upload of some decription ")]
		public string EnclosedTranscript { get; set; }

		public override void Load()
		{
		}

		public DeclarationOfIntentionModel LoadModel()
		{
			DeclarationOfIntentionModel m = new DeclarationOfIntentionModel();

			User user = (User)System.Web.HttpContext.Current.Session["User"];
			long contactNumber = Convert.ToInt64(user.ContactNumber);

			if (contactNumber != null)
			{
				//Get data for this form
				FormData formData = new FormData(FormData.Forms.DeclarationOfIntention);
				List<SelectContactData_CategoriesResult> activityList = formData.GetFormActivities(contactNumber);

				int bandScore;
				Int32.TryParse(activityList.First(activity => activity.ActivityCode == "0TDS").ActivityValueDesc, out bandScore);

				TRFDetails tRFDetails = new TRFDetails
					{
						DateOfTest = activityList.First(activity => activity.ActivityCode == "0TDS").ActivityDate,
						BandScore = bandScore
					};

				m = new DeclarationOfIntentionModel
							 {
								 TakeTestPlanDate = activityList.First(activity => activity.ActivityCode == "0PTD").ActivityDate,
								 TrfDetails = tRFDetails,
								 YearToSit = Convert.ToInt32(activityList.First(activity => activity.ActivityCode == "0YPE").ActivityValueCode)
							 };
			}

			return m;
		}

		public override void Save()
		{
			User user = (User)System.Web.HttpContext.Current.Session["User"];
			UserID = Convert.ToInt64(user.ContactNumber);

			if (UserID == null)
			{
				return;
			}

			//Save activities

			//Year to sit
			Utils.AddActivity(
									UserID,
									"0YPE",
									YearToSit.ToString(),
									"",
									DateTime.Now,
									"WEB"
								);
			// IELTS Activity commit
			Utils.AddActivity(
								 UserID,
								 "0PTD",
									"Y",
									"",
									TakeTestPlanDate,
								 "WEB"
							 );

				//TRF file upload				
				var bytes = new byte[IELTS.InputStream.Length];
				Int64 data = IELTS.InputStream.Read(bytes, 0, Convert.ToInt32(IELTS.InputStream.Length));
				var varBinaryData = Convert.ToBase64String(bytes, 0, Convert.ToInt32(bytes.Length));
				var addCommunicationsLogParameters = new AddCommunicationsLogParameters
					{
						AddresseeContactNumber = user.ContactNumber,
						AddresseeAddressNumber = user.AddressNumber,
						SenderContactNumber = user.ContactNumber,
						SenderAddressNumber = user.AddressNumber,
						Dated = Convert.ToDateTime(DateTime.Now),
						Direction = "U",
						DocumentType = "",
						Topic = "",
						SubTopic = "",
						DocumentClass = "",
						DocumentSubject = "",
						Precis = ""
					};
				var xmlHelper = new XMLHelper(); //XML helper to serialize and deserialize objects
				var serializedParameters = xmlHelper.SerializeToXml(addCommunicationsLogParameters);
				//response = client.AddCommunicationsLog(serializedParameters);

			//TRF details
			Utils.AddActivity(
								UserID,
								"0TDS",
								"A", //TrfDetails.BandScore.ToString(),
								"",
								TrfDetails.DateOfTest,
								"WEB"
							);

			//Vet degree details: Title
			Utils.AddActivity(
								UserID,
								"0TPD",
								"Y",
								PrimaryVetinaryDegree.Name + " (" + PrimaryVetinaryDegree.Abbreviation + ") " + PrimaryVetinaryDegree.Document,
								GraduationDate,
								"WEB"
							);

			//Vet degree details: Normal course length
			Utils.AddActivity(
								UserID,
								"0NLC",
								University.
								DateTime.Now,
								"WEB"
							);

			//Vet degree details: University name
			Utils.AddActivity(
								UserID,
								"0UN",
								University.
								DateTime.Now,
								"WEB"
							);

		}
	}
}
