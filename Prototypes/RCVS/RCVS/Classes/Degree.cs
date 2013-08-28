using System.ComponentModel;
using System.Web;

namespace RCVS.Classes
{
	public class Degree
	{
		public string Name { get; set; }
		public string Abbreviation { get; set; }

		[DisplayName("Upload your veterinary degree here")]
		public HttpPostedFileBase Document { get; set; }

	}
}