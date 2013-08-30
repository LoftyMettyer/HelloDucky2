using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Security.Policy;
using System.Xml.Linq;
using System.Web;
using System.Web.Mvc;
using RCVS.Classes;
using RCVS.Enums;
using RCVS.Helpers;
using RCVS.IRISWebServices;
using RCVS.Structures;
using RCVS.WebServiceClasses;

namespace RCVS.Models
{
	public class DeclarationOfIntentionModel : BaseModel
	{
		public int YearToSit { get; set; } //To hold the value of the selected year to sit

		[DisplayName("Select the year in which you plan to sit the statutory membership examination")]
		public IEnumerable<SelectListItem> YearsDropdown { get; set; }

		public string Activity { get; set; }
		public IEnumerable<SelectListItem> Activities { get; set; }

		[DisplayName("Do you plan to 'see practice'?")]
		public string PlanToSeePractice { get; set; }

		[DisplayName("Are you currently seeing practice or have you made arrangements?")]
		public string CurrentlySeeingPractice { get; set; }

		public List<PracticeArrangement> PracticeArrangements { get; set; }

		[DisplayName("Upload your IELTS test report form here")]
		public HttpPostedFileBase IELTS { get; set; }

		[DisplayName("When do you plan to take the test?")]
		public DateTime? TakeTestPlanDate { get; set; }

		[DisplayName("If you have taken a test, give details and send your TRF for verification")]
		public TRFDetails TrfDetails { get; set; }

		[DisplayName("Title of primary veterinary degree")]
		public Degree PrimaryVeterinaryDegree { get; set; }

		public University UniversityThatAwardedDegree { get; set; }

		[DisplayName("Date of graduation")]
		public DateTime? GraduationDate { get; set; }

		[DisplayName("When did you start your course?")]
		public DateTime? CourseStartDate { get; set; }

		[DisplayName("When did you complete your course?")]
		public DateTime? CourseEndDate { get; set; }

		public int NormalCourseLength { get; set; } //To hold the value of the selected normal course length

		[DisplayName("What is the normal length of your course?")]
		public IEnumerable<SelectListItem> NormalCourseLengthDropdown { get; set; }

		[DisplayName("Have you enclosed a transcript?")]
		public bool? HasEnclosedTranscript { get; set; }

		[DisplayName("Upload English transcript here")]
		public HttpPostedFileBase EnclosedTranscript { get; set; }

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
				//Get data for this form and user
				FormData formData = new FormData(FormData.Forms.DeclarationOfIntention, contactNumber);
				List<SelectContactData_CategoriesResult> activityList = formData.GetFormActivities();

				TRFDetails tRFDetails = new TRFDetails();
				var primaryVeterinaryDegree = new Degree();
				var universityThatAwardedDegree = new University();
				var practiceArrangements  = new List<PracticeArrangement>();

				//See Practice list...

				string response;

				var client = new IRISWebServices.NDataAccessSoapClient(); //Client to call the web services
				var xmlHelper = new XMLHelper(); //XML helper to serialize and deserialize objects

				//Get Position
				var selectContactDataParameters = new SelectContactDataParameters() { ContactNumber = contactNumber };
				var serializedParameters = xmlHelper.SerializeToXml(selectContactDataParameters); //Serialize to XML to pass to the web services

				response = client.SelectContactData(IRISWebServices.XMLContactDataSelectionTypes.xcdtContactPositions, serializedParameters);

				var doc = XDocument.Parse(response);

				var query = from data in doc.Descendants("DataRow")
										select new PracticeArrangement
										{
											PracticeName = (string)data.Element("ContactName"),
											CurrentOrPlanned = ((string)data.Element("PositionSeniority")=="P"?  CurrentOrPlanned.Planned: CurrentOrPlanned.Current),
											StartDate = DateTime.ParseExact((string)data.Element("ValidFrom"), "dd/MM/yyyy", null),
											EndDate = DateTime.ParseExact((string)data.Element("ValidTo"), "dd/MM/yyyy", null),
											VetName = (string)data.Element("Position")											
										};

				practiceArrangements = query.ToList();				

				if (Utils.ActivityIndex(activityList, "0TDS") >= 0)
				{
					int bandScore;
					Int32.TryParse(activityList.First(activity => activity.ActivityCode == "0TDS").ActivityValueCode, out bandScore);

					tRFDetails = new TRFDetails
						{
							BandScore = bandScore
						};
					if (!activityList.First(activity => activity.ActivityCode == "0TDS").ActivityDate.Equals(DateTime.MinValue))
					{
						tRFDetails.DateOfTest = activityList.First(activity => activity.ActivityCode == "0TDS").ActivityDate;
					}
				}

				if (Utils.ActivityIndex(activityList, "0TPD") >= 0)
				{
					primaryVeterinaryDegree = new Degree
						{
							Name = activityList.First(activity => activity.ActivityCode == "0TPD").Notes
						};
				}

				universityThatAwardedDegree = new University
					{
						Name =
							(Utils.ActivityIndex(activityList, "0UN") >= 0)
								? activityList.First(activity => activity.ActivityCode == "0UN").Notes
								: "",
						City =
							(Utils.ActivityIndex(activityList, "0UCC") >= 0)
								? activityList.First(activity => activity.ActivityCode == "0UCC").Notes
								: "",
						Country =
							(Utils.ActivityIndex(activityList, "0UC") >= 0)
								? activityList.First(activity => activity.ActivityCode == "0UC").ActivityValueCode
								: ""
					};

				m = new DeclarationOfIntentionModel
					{
						TrfDetails = tRFDetails,
						PrimaryVeterinaryDegree = primaryVeterinaryDegree,
						UniversityThatAwardedDegree = universityThatAwardedDegree,
						PracticeArrangements = practiceArrangements
					};


				if (Utils.ActivityIndex(activityList, "0PTD") >= 0)
				{
					m.TakeTestPlanDate = activityList.First(activity => activity.ActivityCode == "0PTD").ActivityDate;
				}

				if (Utils.ActivityIndex(activityList, "0YPE") >= 0)
				{
					m.YearToSit = Convert.ToInt32(activityList.First(activity => activity.ActivityCode == "0YPE").ActivityValueCode);
				}

				if (Utils.ActivityIndex(activityList, "0TPD") >= 0)
				{
					m.GraduationDate = activityList.First(activity => activity.ActivityCode == "0TPD").ActivityDate;
				}
				if (Utils.ActivityIndex(activityList, "0NLC") >= 0)
				{
					m.NormalCourseLength = Convert.ToInt32(activityList.First(activity => activity.ActivityCode == "0NLC").ActivityValueCode);
				}
				if (Utils.ActivityIndex(activityList, "0PSP") >= 0)
				{
					m.PlanToSeePractice = activityList.First(activity => activity.ActivityCode == "0PSP").ActivityValueCode;
				}
				if (Utils.ActivityIndex(activityList, "0CSP") >= 0)
				{
					m.CurrentlySeeingPractice = activityList.First(activity => activity.ActivityCode == "0CSP").ActivityValueCode;
				}
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
									(TakeTestPlanDate.HasValue) ? (DateTime)TakeTestPlanDate : DateTime.MinValue,
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
					Utils.LogWebServiceCall("GetLookupData", "NONE", response); //Log the call and response
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
					Utils.LogWebServiceCall("AddCommunicationsLog", serializedParameters, response); //Log the call and response

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
						Utils.LogWebServiceCall("UpdateDocumentFile", serializedParameters, response); //Log the call and response
					}

					client.Close();
				}
			}


			//TRF details
			Utils.AddActivity(
								UserID,
								"0TDS",
								TrfDetails.BandScore.ToString(),
								"",
								(TrfDetails.DateOfTest.HasValue) ? (DateTime)TrfDetails.DateOfTest : DateTime.MinValue,
								"WEB"
							);

			//Vet degree details: Title and graduation date
			Utils.AddActivity(
								UserID,
								"0TPD",
								"Y",
								PrimaryVeterinaryDegree.Name,
								(GraduationDate.HasValue) ? (DateTime)GraduationDate : DateTime.MinValue,
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

			//Do you plan to see practice?
			Utils.AddActivity(
								UserID,
								"0PSP",
								PlanToSeePractice,
								"",
								DateTime.Now,
								"WEB"
							);

			//Are you currentl seeing practice?
			Utils.AddActivity(
								UserID,
								"0CSP",
								CurrentlySeeingPractice,
								"",
								DateTime.Now,
								"WEB"
							);
		}
	}
}
