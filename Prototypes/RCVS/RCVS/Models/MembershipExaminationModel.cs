using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using RCVS.Classes;

namespace RCVS.Models
{
	public class MembershipExaminationModel
	{

		[DisplayName("Reason for renewal of intention to sit?")]
		public string RenewalReason { get; set; }

		public IEnumerable<SelectListItem> RenewalReasons { get; set; }

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

		[ DisplayName("If you have taken a test, give details")]
		public TRFDetails PreviousTest { get; set; }


		public void Save()
		{
			int _Save = 1;


		}



	}
}