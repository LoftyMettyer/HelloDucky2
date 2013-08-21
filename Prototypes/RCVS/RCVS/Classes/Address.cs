using System.ComponentModel.DataAnnotations;

namespace RCVS.Classes
{
	public class Address
	{
		public string AddressLine1 { get; set; }
		public string AddressLine2 { get; set; }
		public string AddressLine3 { get; set; }
		public string Town { get; set; }
		public string County { get; set; }
		public string Country { get; set; }

		[Required]
		public string Postcode { get; set; }
	}
}