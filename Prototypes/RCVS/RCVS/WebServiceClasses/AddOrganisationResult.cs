using System.Xml.Serialization;

namespace RCVS.WebServiceClasses
{
	[XmlRoot("Result")]
	public class AddOrganisationResult
	{
		public long ContactNumber { get; set; }
		public long AddressNumber { get; set; }
	}
}