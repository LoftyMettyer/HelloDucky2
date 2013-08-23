using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;



namespace RCVS.WebServiceClasses
{
	[System.Xml.Serialization.XmlRoot("Result")]
	public class SelectContactDataResult
	{
		//public TYPE Type { get; set; }
		public long ContactNumber { get; set; }
		public string ActivityCode { get; set; }
		public string ActivityValueCode { get; set; }
		public int Quantity { get; set; } //int??
		public DateTime ActivityDate { get; set; }
		public string SourceCode { get; set; }
		public DateTime ValidFrom { get; set; }
		public DateTime ValidTo { get; set; }
		public string AmendedBy { get; set; }
		public DateTime AmendedOn { get; set; }	//datetime??
		public string Notes { get; set; }
		public string ActivityDesc { get; set; }
		public string ActivityValueDesc { get; set; }
		public string SourceDesc { get; set; }
		public string RgbActivityValue { get; set; }
		public string NoteFlag { get; set; }
		public string Status { get; set; }
		public string Access { get; set; }
		public string StatusOrder { get; set; }
	}
}