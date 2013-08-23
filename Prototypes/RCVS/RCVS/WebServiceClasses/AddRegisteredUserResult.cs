using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace RCVS.WebServiceClasses
{
	[System.Xml.Serialization.XmlRoot("Result")]
	public class AddRegisteredUserResult
	{
		public string UserName { get; set; }
		public string Password { get; set; }
	}
}