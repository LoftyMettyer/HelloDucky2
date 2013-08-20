using System;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;

namespace RCVS.Models
{
	public class Form1A2
	{
		[Required]
		[DisplayName("YEAR in which you plan to sit the statutory membership examination")]
		public int YearToSit { get; set; }

		public string Footnote1
		{
			get { return "If you do not sit this examination at the next available session or, if you must re-confirm your intention to sit within a reasonable period of time by completing a renewal of intention form"; }
		}

		[DisplayName("Do you plan to 'see practice'?")]
		public bool? PlanToSeePractice { get; set; }

		[DisplayName("Are you currently seeing practice or have you made arrangements?")]
		public bool? CurrentlySeeingPractice { get; set; }

		public string IELTS { get; set; }

		public DateTime? TakeTestPlanDate { get; set; }

//		public TRF


		public Classes.University UniversityAwarded { get; set; }



	}
}