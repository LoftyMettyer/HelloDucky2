using System;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;

namespace RCVS.Models
{
	public class Form1A2Model
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

		[DisplayName("TODO - This need to be a file upload of some decription ")]
		public string IELTS { get; set; }

		[DisplayName("When do you plan to take the test?")]
		public DateTime? TakeTestPlanDate { get; set; }

		[DisplayName("If you have taken a test, give details and send your TRF for verification")]
		public Classes.TRFDetails TrfDetails { get; set; }

		[DisplayName("Title of primary veterinary degree and recognised abbreviation if any")]
		public Classes.Degree PrimaryVetinaryDegree { get; set; }

		public Classes.University UniversityAwarded { get; set; }

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
	}
}