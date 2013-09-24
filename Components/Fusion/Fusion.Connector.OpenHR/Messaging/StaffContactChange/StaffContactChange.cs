using System;
using System.Xml.Serialization;
using Fusion.Connector.OpenHR.MessageComponents.Component;
using Fusion.Connector.OpenHR.MessageComponents.Data;
using Fusion.Connector.OpenHR.MessageComponents.Enums;
using Fusion.Connector.OpenHR.Messaging.Base;

namespace Fusion.Connector.OpenHR.MessageComponents
{
	[System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "4.0.30319.17929")]
	[SerializableAttribute]
	[XmlType(AnonymousType = true, Namespace = "http://advancedcomputersoftware.com/xml/fusion/socialCare")]
	[XmlRoot(Namespace = "http://advancedcomputersoftware.com/xml/fusion/socialCare", IsNullable = false, ElementName = "staffContactChange")]
	public class StaffContactChange : BaseMessageComponent
	{

		public StaffContactChange()
		{
			data = new StaffContactChangeData();
		}

		public StaffContactChange(Guid busRef, Guid? parentRef, Contact contact)
		{
			staffContactRef = busRef.ToString();
			staffRef = parentRef.ToString();
			data = new StaffContactChangeData
			{
				staffContact = contact,
				recordStatus = contact.isRecordInactive == true ? RecordStatusStandard.Inactive : RecordStatusStandard.Active,
				auditUserName = "OpenHR User"
			};
		}

		public StaffContactChangeData data { get; set; }

		[XmlAttributeAttribute]
		public string staffContactRef { get; set; }

		[XmlAttributeAttribute]
		public string staffRef { get; set; }
	}


}
