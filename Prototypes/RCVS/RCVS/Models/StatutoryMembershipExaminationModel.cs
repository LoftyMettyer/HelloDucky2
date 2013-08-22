using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using RCVS.Classes;
using RCVS.Structures;

namespace RCVS.Models
{
	public class StatutoryMembershipExaminationModel : BaseModel
	{
		[Required]
		[DisplayName("Are you a new applicant?")]
		public bool IsNewApplicant { get; set; }

		[DisplayName("Year of last application?")]
	public int YearOfLastApplication { get; set; }

	[Required]
	[DisplayName("Please select the subjects you have permission to sit")]
	public List<Subject> SubjectsWithPermission { get; set; }

	public List<Qualification> Qualifications { get; set; }

	public List<Employment> EmploymentHistory { get; set; }

	[Required]
	[DisplayName("Are you, or have you been at any time, in the Register of persons qualified to practise veterinary surgery in any country or state?")]
	public bool PreviouslyRegistered { get; set; }

	public string RegistrationAuthority { get; set; }
	public Address RegistrationAuthorityAddress { get; set; }

	[DisplayName("Date of registration")]
	public System.DateTime RegistrationDate { get; set; }

	[DisplayName("Registration expiry date")]
	public System.DateTime RegistrationExpiryDate { get; set; }

	[Required]
	[DisplayName("Have you been banned or suspended at any time from practising or refused permission to practise veterinary surgery in any country or state?")]
	public bool PreviouslyBanned { get; set; }

	[DisplayName("Please give reason(s) for your ban or suspension")]
	public string BanReasons { get; set; }

	[DisplayName("If you are not currently registered or if you have never been registered to practise in any country, please explain why.")]
	public string NotRegisteredReasons { get; set; }

	public ExamPayment ExaminationFees { get; set; }

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
		}

		public override void Save()
		{
			throw new System.NotImplementedException();
		}
	}
}