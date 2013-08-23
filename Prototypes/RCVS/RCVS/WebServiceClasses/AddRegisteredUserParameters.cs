using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace RCVS.WebServiceClasses
{
	public class AddRegisteredUserParameters
	{
		public long ContactNumber { get; set; }
		public string UserName { get; set; }
		public string Password { get; set; }
		public string EMailAddress { get; set; }
	}
}