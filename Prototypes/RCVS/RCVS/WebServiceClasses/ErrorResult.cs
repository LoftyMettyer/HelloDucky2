using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace RCVS.WebServiceClasses
{
	[System.Xml.Serialization.XmlRoot("Result")]
	public class ErrorResult
	{
		public string ErrorMessage { get; set; }
		public string Source { get; set; }
		public int ErrorNumber { get; set; }
		public string Module { get; set; }
		public string Method { get; set; }
		public string StackTrace { get; set; }
	}
}