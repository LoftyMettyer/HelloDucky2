using System.ComponentModel;
using System.Web.Mvc;

namespace RCVS.Classes
{
	public class RenewalReason : SelectListItem
	{
		[DisplayName("Value")]
		public string Reason { get; set; }

		[DisplayName("Text")]
		public string Description { get; set; }
		public string Automatic { get; set; }
	}
}