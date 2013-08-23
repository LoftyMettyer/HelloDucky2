using System.Xml.Serialization;

namespace RCVS.WebServiceClasses
{
	[XmlRoot("Result")]
	public class AddPositionResult
	{
		public long ContactPositionNumber { get; set; }
	}
}