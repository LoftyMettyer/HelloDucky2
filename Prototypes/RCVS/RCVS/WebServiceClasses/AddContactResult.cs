using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace RCVS.WebServiceClasses
{
	[System.Xml.Serialization.XmlRoot("Result")]
	public class AddContactResult
	{
		public long ContactNumber { get; set; }
		public long AddressNumber { get; set; }
		public long ContactPositionNumber { get; set; }
		public string ActivityGroup { get; set; }
	}
}