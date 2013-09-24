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
	[XmlRoot(Namespace = "http://advancedcomputersoftware.com/xml/fusion/socialCare", IsNullable = false, ElementName = "staffContractChange")]
	public class StaffContractChange : BaseMessageComponent
	{

		public StaffContractChange() {}

		public StaffContractChange(Guid busRef, Guid? parentRef, Contract contract)
		{
			staffContractRef = busRef.ToString();
			staffRef = parentRef.ToString();
			data = new StaffContractChangeData
			{
				staffContract = contract,
				recordStatus = contract.isRecordInactive == true ? RecordStatusStandard.Inactive : RecordStatusStandard.Active,
				auditUserName = "OpenHR User"
			};
		}

		public StaffContractChangeData data { get; set; }

		[XmlAttributeAttribute]
		public string staffContractRef { get; set; }

		[XmlAttributeAttribute]
		public string staffRef { get; set; }
	}

}
