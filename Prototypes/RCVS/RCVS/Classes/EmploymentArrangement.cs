using System;
using System.ComponentModel;

namespace RCVS.Classes
{
	public class EmploymentArrangement
	{
		[DisplayName("Indicate Current or Planned")]
		public Enums.CurrentOrPlanned CurrentOrPlanned { get; set; }

		[DisplayName("Veterinary practice or other establishment")]
		public string PracticeName { get; set; }

		public Address Address { get; set; }

		[DisplayName("Position held")]
		public string Position { get; set; }

		[DisplayName("Start date")]
		public DateTime? StartDate { get; set; }

		[DisplayName("End date")]
		public DateTime? EndDate { get; set; }

	}
}