using System;
using System.ComponentModel;

namespace RCVS.Classes
{
	public class PracticeArrangement
	{
		public Enums.CurrentOrPlanned CurrentOrPlanned { get; set; }

		[DisplayName("Veterinary practice or other establishment")]
		public string PracticeName { get; set; }

		public Address Address { get; set; }

		[DisplayName("Full name of supervising vet")]
		public string VetName { get; set; }

		public DateTime? StartDate { get; set; }
		public DateTime? EndDate { get; set; }

	}
}