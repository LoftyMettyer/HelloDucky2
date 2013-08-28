using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using System.Web;
using System.Web.Mvc;
using RCVS.Classes;
using RCVS.Helpers;
using RCVS.IRISWebServices;
using RCVS.Structures;
using RCVS.WebServiceClasses;

namespace RCVS.Models
{
	public class DeclarationOfIntentionModel : BaseModel
	{
		public int YearToSit { get; set; } //To hold the value of the selected year to sit

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

		public List<PracticeArrangement> PracticeArrangements { get; set; }

		[DisplayName("Upload your IELTS test report form here")]
		public HttpPostedFileBase IELTS { get; set; }

		[DisplayName("When do you plan to take the test?")]
		public DateTime TakeTestPlanDate { get; set; }

		[DisplayName("If you have taken a test, give details and send your TRF for verification")]
		public TRFDetails TrfDetails { get; set; }

		[DisplayName("Title of primary veterinary degree and recognised abbreviation if any")]
		public Degree PrimaryVeterinaryDegree { get; set; }

		public University UniversityThatAwardedDegree { get; set; }

		[DisplayName("Date of graduation")]
		public DateTime? GraduationDate { get; set; }

		[DisplayName("When did you start your course?")]
		public DateTime? CourseStartDate { get; set; }

		[DisplayName("When did you complete your course?")]
		public DateTime? CourseEndDate { get; set; }

		public int NormalCourseLength { get; set; } //To hold the value of the selected normal course length

		[Required]
		[DisplayName("What is the normal length of your course?")]
		public IEnumerable<SelectListItem> NormalCourseLengthDropdown { get; set; }

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
				Int32.TryParse(activityList.First(activity => activity.ActivityCode == "0TDS").ActivityValueCode, out bandScore);

				TRFDetails tRFDetails = new TRFDetails
					{
						DateOfTest = activityList.First(activity => activity.ActivityCode == "0TDS").ActivityDate,
						BandScore = bandScore
					};

				Degree primaryVeterinaryDegree = new Degree
					{
						Name = activityList.First(activity => activity.ActivityCode == "0TPD").Notes
					};

				University universityThatAwardedDegree = new University
					{
						Name = activityList.First(activity => activity.ActivityCode == "0UN").Notes,
						City = activityList.First(activity => activity.ActivityCode == "0UCC").Notes,
						Country = activityList.First(activity => activity.ActivityCode == "0UC").ActivityValueCode
					};

				m = new DeclarationOfIntentionModel
							 {
								 TakeTestPlanDate = activityList.First(activity => activity.ActivityCode == "0PTD").ActivityDate,
								 TrfDetails = tRFDetails,
								 YearToSit = Convert.ToInt32(activityList.First(activity => activity.ActivityCode == "0YPE").ActivityValueCode),
								 PrimaryVeterinaryDegree = primaryVeterinaryDegree,
								 GraduationDate = activityList.First(activity => activity.ActivityCode == "0TPD").ActivityDate,
								 NormalCourseLength = Convert.ToInt32(activityList.First(activity => activity.ActivityCode == "0NLC").ActivityValueCode),
								 UniversityThatAwardedDegree = universityThatAwardedDegree
							 };
			}

			return m;
		}

		public override void Save()
		{
			User user = (User)System.Web.HttpContext.Current.Session["User"];
			long UserID = Convert.ToInt64(user.ContactNumber);

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
			//Done in three parts:
			//1. get a document number
			//2. get the document application type
			//3. upload via the document number.
			if (IELTS != null)
			{
				var extension = Path.GetExtension(IELTS.FileName);

				if (extension != null)
				{
					extension = extension.ToUpper();
					//get package type for this file type - no package type, no upload!
					const XMLLookupDataTypes lookupDataType = XMLLookupDataTypes.xldtPackages;
					var client = new NDataAccessSoapClient();
					var response = client.GetLookupData(lookupDataType, "");
					var package = "";

					var doc = XDocument.Parse(response);

					foreach (XElement xe in doc.Descendants("DataRow"))
					{
						var element = xe.Element("DocfileExtension");
						if (element != null)
						{
							var docFileExtension = element.Value.ToUpper();
							if (docFileExtension == extension)
							{
								package = xe.Element("Package").Value;
								break;
							}
						}
					}

				var bytes = new byte[IELTS.InputStream.Length];
					IELTS.InputStream.Read(bytes, 0, Convert.ToInt32(IELTS.InputStream.Length));
				var addCommunicationsLogParameters = new AddCommunicationsLogParameters
					{
						AddresseeContactNumber = user.ContactNumber,
						AddresseeAddressNumber = user.AddressNumber,
						SenderContactNumber = user.ContactNumber,
						SenderAddressNumber = user.AddressNumber,
						Dated = Convert.ToDateTime(DateTime.Now),
							Direction = "I",
							DocumentType = "OTHE",
							Topic = "GEN",
							SubTopic = "CORR",
							DocumentClass = "U",
							DocumentSubject = "TRF Upload",
							Precis = "TRF Upload",
							Package = package
					};
				var xmlHelper = new XMLHelper(); //XML helper to serialize and deserialize objects
				var serializedParameters = xmlHelper.SerializeToXml(addCommunicationsLogParameters);
					client = new NDataAccessSoapClient();
					response = client.AddCommunicationsLog(serializedParameters);

					var xElement = XDocument.Parse(response).Element("Result");
					if (xElement != null)
					{
						var documentNumber = Convert.ToInt32(xElement.Value);
						var updateDocumentFileParameters = new UpdateDocumentFileParameters
							{
								DocumentNumber = documentNumber
							};
						xmlHelper = new XMLHelper();
						serializedParameters = xmlHelper.SerializeToXml(updateDocumentFileParameters);
						response = client.UpdateDocumentFile(serializedParameters, bytes);
					}

					client.Close();
				}
			}


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
								PrimaryVeterinaryDegree.Name,
								(DateTime)GraduationDate,
								"WEB"
							);

			//Vet degree details: Normal course length
			Utils.AddActivity(
								UserID,
								"0NLC",
								NormalCourseLength.ToString(),
								"",
								DateTime.Now,
								"WEB"
							);

			//Vet degree details: University name
			Utils.AddActivity(
								UserID,
								"0UN",
								"Y",
								UniversityThatAwardedDegree.Name,
								DateTime.Now,
								"WEB"
							);

			//Vet degree details: University city
			Utils.AddActivity(
								UserID,
								"0UCC",
								"Y",
								UniversityThatAwardedDegree.City,
								DateTime.Now,
								"WEB"
							);

			//Vet degree details: University country
			Utils.AddActivity(
								UserID,
								"0UC",
								UniversityThatAwardedDegree.Country,
								"",
								DateTime.Now,
								"WEB"
							);
		}
	}
}
