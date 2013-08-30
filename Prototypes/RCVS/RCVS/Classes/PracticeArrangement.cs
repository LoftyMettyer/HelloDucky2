using System;
using System.ComponentModel;

namespace RCVS.Classes
{
	public class PracticeArrangement
	{
		[DisplayName("Indicate Current or Planned")]
		public Enums.CurrentOrPlanned CurrentOrPlanned { get; set; }

		[DisplayName("Veterinary practice or other establishment")]
		public string PracticeName { get; set; }

		public Address Address { get; set; }

		[DisplayName("Full name of supervising vet")]
		public string VetName { get; set; }

		[DisplayName("Start date")]
		public DateTime? StartDate { get; set; }

		[DisplayName("End date")]
		public DateTime? EndDate { get; set; }

	}
}