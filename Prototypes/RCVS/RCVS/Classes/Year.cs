using System.ComponentModel;
using System.Web.Mvc;

namespace RCVS.Classes
{
	public class Year : SelectListItem
	{
		[DisplayName("Value")]
		public string _Value { get; set; }

		[DisplayName("Text")]
		public string _Text { get; set; }
	}
}
