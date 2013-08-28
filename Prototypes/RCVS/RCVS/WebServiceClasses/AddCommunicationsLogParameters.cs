using System;

namespace RCVS.WebServiceClasses
{
	public class AddCommunicationsLogParameters
	{
		public long AddresseeContactNumber { get; set; }
		public long AddresseeAddressNumber { get; set; }
		public long SenderContactNumber { get; set; }
		public long SenderAddressNumber { get; set; }
		public DateTime Dated { get; set; }
		public string Direction { get; set; }
		public string DocumentType { get; set; }
		public string Topic { get; set; }
		public string SubTopic { get; set; }
		public string DocumentClass { get; set; }
		public string DocumentSubject { get; set; }
		public string Precis { get; set; }
		public string Package { get; set; }
	}
}