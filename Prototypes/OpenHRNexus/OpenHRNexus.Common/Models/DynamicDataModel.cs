using System.ComponentModel.DataAnnotations.Schema;
using System.Data;

namespace OpenHRNexus.Common.Models
{
	public class DynamicDataModel
	{
		public int Id { get; set; }

		//[Column("Forename")]
		public string Column1 { get; set; }

		//[Column("Surname")]
		public string Column2 { get; set; }

		//[Column("Title")]
		public string Column3 { get; set; }

		//[Column("Date Of Birth")]
		public string Column4 { get; set; } 
	}
}
