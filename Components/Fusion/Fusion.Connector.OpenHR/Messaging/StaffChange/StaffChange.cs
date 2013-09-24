using System;
using System.CodeDom.Compiler;
using Fusion.Connector.OpenHR.Configuration;
using Fusion.Connector.OpenHR.MessageComponents.Component;
using Fusion.Connector.OpenHR.MessageComponents.Data;
using System.Xml.Serialization;
using Fusion.Connector.OpenHR.MessageComponents.Enums;
using Fusion.Connector.OpenHR.Messaging.Base;
using StructureMap.Attributes;

namespace Fusion.Connector.OpenHR.MessageComponents
{
	[GeneratedCode("xsd", "4.0.30319.17929")]
	[Serializable]
	[XmlType(AnonymousType = true, Namespace = "http://advancedcomputersoftware.com/xml/fusion/socialCare")]
	[XmlRoot(Namespace = "http://advancedcomputersoftware.com/xml/fusion/socialCare", IsNullable = false, ElementName = "staffChange")]
	public class StaffChange : BaseMessageComponent
	{

		public StaffChange() {}

		public StaffChange(Guid BusRef, Staff staff)
		{
			staffRef = BusRef.ToString();
			data = new StaffChangeData {
				staff = staff,
				recordStatus = staff.isRecordInactive == true ? RecordStatusStandard.Inactive : RecordStatusStandard.Active,
				auditUserName = "OpenHR User"
			};
		}

		public StaffChangeData data { get; set; }

		[XmlAttribute]
		public string staffRef { get; set; }
	}


}
