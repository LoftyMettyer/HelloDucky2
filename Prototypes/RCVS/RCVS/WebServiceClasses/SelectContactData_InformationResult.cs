using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace RCVS.WebServiceClasses
{
	[System.Xml.Serialization.XmlRoot("Result")]
	public class SelectContactData_InformationResult
	{
		public long ContactNumber { get; set; }
		public string Postcode { get; set; }
	}
}