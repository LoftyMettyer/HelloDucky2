using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace RCVS.Classes
{
	public class User
	{
		public long ContactNumber { get; set; }
		public long AddressNumber { get; set; }
		public string ContactName { get; set; }
		public string Title { get; set; }
		public string Initials { get; set; }
		public string Forenames { get; set; }
		public string Surname { get; set; }
		public string Honorifics { get; set; }
		public string Salutation { get; set; }
		public string LabelName { get; set; }
	}
}