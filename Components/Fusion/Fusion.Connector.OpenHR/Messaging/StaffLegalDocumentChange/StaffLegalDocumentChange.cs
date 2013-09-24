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
	[XmlRoot(Namespace = "http://advancedcomputersoftware.com/xml/fusion/socialCare", IsNullable = false, ElementName = "staffLegalDocumentChange")]
	public class StaffLegalDocumentChange : BaseMessageComponent
	{
		public StaffLegalDocumentChange() { }

		public StaffLegalDocumentChange(Guid busRef, Guid? parentRef, LegalDocument legalDocument)
		{
			staffLegalDocumentRef = busRef.ToString();
			staffRef = parentRef.ToString();
			data = new StaffLegalDocumentChangeData
			{
				staffLegalDocument = legalDocument,
				recordStatus = legalDocument.isRecordInactive == true ? RecordStatusStandard.Inactive : RecordStatusStandard.Active,
				auditUserName = "OpenHR User"
			};
		}

		public StaffLegalDocumentChangeData data { get; set; }

		[XmlAttributeAttribute]
		public string staffLegalDocumentRef { get; set; }

		[XmlAttributeAttribute]
		public string staffRef { get; set; }
	}

}
