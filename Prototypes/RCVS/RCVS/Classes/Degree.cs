using System.ComponentModel;

namespace RCVS.Classes
{
	public class Degree
	{
		public string Name { get; set; }
		public string Abbreviation { get; set; }

		[DisplayName("TODO - This need to be a file upload of some decription ")]
		public string Document { get; set; }

	}
}