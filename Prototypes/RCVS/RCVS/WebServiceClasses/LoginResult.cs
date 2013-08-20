using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace RCVS.WebServiceClasses
{
	[System.Xml.Serialization.XmlRoot("Result")]
	public class LoginResult
	{
		public long ContactNumber { get; set; }
		public string UserLogname { get; set; }
		public string UserDepartment { get; set; }
		public string DatabaseDescription { get; set; }
		public long AddressNumber { get; set; }
	}
}