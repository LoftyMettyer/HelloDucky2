using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace RCVS.WebServiceClasses
{
	[System.Xml.Serialization.XmlRoot("Result")]
	public class LoginRegisteredUserResult
	{
		public long ContactNumber { get; set; }
		public string Password { get; set; }
		public string EMailAddress { get; set; }
		public string SecurityQuestion { get; set; }
		public string SecurityAnswer { get; set; }
		public long AddressNumber { get; set; }
		public string UserLogname { get; set; }
		public string UserDepartment { get; set; }
	}
}