using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Web.Mvc;

namespace RCVS.Models
{
	public class Form1AModel
	{
		[Required]
		[DisplayName("Your name in full including your title")]
		public string YourNameInFull { get; set; }

		[Required]
		[DisplayName("All first names")]
		public string AllFirstNames { get; set; }

		[Required]
		[DisplayName("All surnames")]
		public string AllSurnames { get; set; }

		[Required]
		[DisplayName("Day")]
		public int DayOfBirth { get; set; }

		[Required]
		[DisplayName("Month")]
		public int MonthOfBirth { get; set; }

		[Required]
		[DisplayName("Year")]
		public int YearOfBirth { get; set; }

		[Required]
		[DisplayName("Address line 1")]
		public string AddressLine1 { get; set; }

		[Required]
		[DisplayName("Address line 2")]
		public string AddressLine2 { get; set; }

		[DisplayName("Address line 3")]
		public string AddressLine3 { get; set; }

		[Required]
		[DisplayName("Postcode (UK)")]
		public string Postcode { get; set; }
	}
}
