using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace RCVS.WebServiceClasses
{
	[System.Xml.Serialization.XmlRoot("Result")]
	public class AddActionFromTemplateResult
	{
		public long ActionNumber { get; set; }
	}
}