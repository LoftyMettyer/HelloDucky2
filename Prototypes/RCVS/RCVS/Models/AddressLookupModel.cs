using System.ComponentModel;
using System.ComponentModel.DataAnnotations;

namespace RCVS.Models
{
	public class AddressLookupModel
	{
		[Required]
		[DisplayName("Please enter your postcode")]
		public string Postcode { get; set; }
	}
}