using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace RCVS.WebServiceClasses
{
	public class AddActivityWithExtraParametersParameters
	{
		public long ContactNumber { get; set; }
		public string Activity { get; set; }
		public string ActivityValue { get; set; }
		public string Source { get; set; }
		public DateTime? ActivityDate { get; set; }
		public DateTime? ValidFrom { get; set; }
		public DateTime? ValidTo  { get; set; }
		public string Notes { get; set; }
	}
}