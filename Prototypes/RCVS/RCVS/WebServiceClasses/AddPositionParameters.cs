using System;

namespace RCVS.WebServiceClasses
{
	public class AddPositionParameters
	{
	public long ContactNumber { get; set; }
	public long OrganisationNumber { get; set; }
	public long AddressNumber { get; set; }
	public string Position { get; set; }
	public DateTime ValidFrom { get; set; }
	public DateTime ValidTo { get; set; }
	public string PositionSeniority { get; set; }
	}
}