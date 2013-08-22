using System;

namespace RCVS.WebServiceClasses
{
	public class AddActivityParameters
	{
		public long ContactNumber { get; set; }
		public string Activity { get; set; }
		public string ActivityValue { get; set; }
		public string Source { get; set; }
		public DateTime? ActivityDate { get; set; }		

	}
}