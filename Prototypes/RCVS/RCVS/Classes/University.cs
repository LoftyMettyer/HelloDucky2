using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Web.Mvc;

namespace RCVS.Classes
{
	public class University
	{
		[Required]
		[DisplayName("")]
		public string Name { get; set; }
		[Required]
		public string City { get; set; }

		public string Country { get; set; }

		[Required]
		[DisplayName("Country")]
		public IEnumerable<SelectListItem> CountriesDropdown { get; set; }
	}
}