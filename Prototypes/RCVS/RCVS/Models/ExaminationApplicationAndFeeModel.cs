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
	public class ExaminationApplicationAndFeeModel : BaseModel
	{
		[Required]
		[DisplayName("Are you a new applicant?")]
		public bool IsNewApplicant { get; set; }

		[DisplayName("If no, year of last application?")]
		public string YearOfLastApplication { get; set; } //To hold the value of the selected year
		public List<SelectListItem> YearOfLastApplicationDropDown { get; set; }

		[Required]
		public IEnumerable<SelectListItem> SubjectsWithPermissionDropDown { get; set; }
		[DisplayName("Please select the subject you have permission to sit")]
		public string SubjectWithPermission { get; set; } //To hold the value of the selected subject

		public List<Qualification> Qualifications { get; set; }

		public List<Employment> EmploymentHistory { get; set; }

		[Required]
		[DisplayName("Are you, or have you been at any time, in the Register of persons qualified to practise veterinary surgery in any country or state?")]
		public string PreviouslyRegistered { get; set; }

		[DisplayName("Registration authority name and address")]
		[DataType(DataType.MultilineText)]
		public string RegistrationAuthority { get; set; }

		[DisplayName("Date of registration")]
		public DateTime? RegistrationDate { get; set; }

		[DisplayName("Registration expiry date")]
		public DateTime? RegistrationExpiryDate { get; set; }

		[Required]
		[DisplayName("Have you been banned or suspended at any time from practising or refused permission to practise veterinary surgery in any country or state?")]
		public string PreviouslyBanned { get; set; }

		[DisplayName("If Yes, please give reasons")]
		[DataType(DataType.MultilineText)]
		public string BanReasons { get; set; }

		[DisplayName("If you are not currently registered or if you have never been registered to practise in any country, please explain why.")]
		[DataType(DataType.MultilineText)]
		public string NotRegisteredReasons { get; set; }

		[DisplayName("The amount you are paying")]
		public string AmountYouArePaying { get; set; } //To hold the value of the selected amount
		public List<SelectListItem> AmountYouArePayingDropDown { get; set; }

		[DisplayName("Please confirm with your name and date")]
		public Confirmation FeeConfirmation { get; set; }

		public override void Load()
		{
			Qualifications = new List<Qualification>();
			Qualifications.Add(new Qualification
				{
					AwardingBody = "Staffordshire University",
					Name = "Software Science",
					ObtainedDate = System.DateTime.Now
				});

			// Retrieve from web service
			string response;
			var client = new IRISWebServices.NDataAccessSoapClient();

			var XmlHelper = new XMLHelper();
			//var addActivityParameters = new FindActions() { UserID = "571", myActions = "0PSP" };
			//var serializedParameters = XmlHelper.SerializeToXml(addActivityParameters);

			//response = client.FindActions(serializedParameters);
			//AddActivity(serializedParameters);

			var addParameters = new GetLookupDataParameters() { UserID = "571", Activity = "0PSP", ContactGroup = "", OrganisationGroup = "", Product = "", Topic = "" };
			var serializedParameters2 = XmlHelper.SerializeToXml(addParameters);

			var lookupDataType = IRISWebServices.XMLLookupDataTypes.xldtActivitiesAndValues;
			response = client.GetLookupData(lookupDataType, serializedParameters2);
			//AddActivity(serializedParameters);







			EmploymentHistory = new List<Employment>();
			EmploymentHistory.Add(new Employment { City = "Aberdare", Country = "Wales", FromDate = System.DateTime.Now.AddYears(-3), ToDate = System.DateTime.Now.AddYears(-2), Position = "Junior Vet", PracticeName = "Cows & Sons" });
			EmploymentHistory.Add(new Employment { City = "Guildford", Country = "England", FromDate = System.DateTime.Now.AddYears(-2), ToDate = System.DateTime.Now.AddYears(-1), Position = "Chief Vet", PracticeName = "Horse Bros" });

		}

		public override void Save()
		{
		}
	}
}