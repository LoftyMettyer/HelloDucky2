using System.Xml.Serialization;

namespace RCVS.WebServiceClasses
{
	[XmlRoot("ResultSet")]
	public class FindOrganisationResults
	{
		public string OrganisationNumber { get; set; }
		public string ContactNumber { get; set; }
		public string Name { get; set; }
		public string SortName { get; set; }
		public string Abbreviation { get; set; }
		public string DiallingCode { get; set; }
		public string StdCode { get; set; }
		public string Telephone { get; set; }
		public string Status { get; set; }
		public string OwnershipGroup { get; set; }
		public string AddressNumber { get; set; }
		public string Address { get; set; }
		public string HouseName { get; set; }
		public string Town { get; set; }
		public string County { get; set; }
		public string Postcode { get; set; }
		public string Branch { get; set; }
		public string Country { get; set; }
		public string AddressType { get; set; }
		public string BuildingNumber { get; set; }
		public string StatusDesc { get; set; }
		public string OwnershipGroupDesc { get; set; }
		public string PrincipalDepartmentDesc { get; set; }
		public string OwnershipAccessLevel { get; set; }
		public string OwnershipAccessLevelDesc { get; set; }

	}
}