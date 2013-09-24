using System;
using Fusion.Connector.OpenHR.MessageComponents.Component;
using Fusion.Connector.OpenHR.MessageComponents.Data;
using System.Xml.Serialization;
using Fusion.Connector.OpenHR.MessageComponents.Enums;
using Fusion.Connector.OpenHR.Messaging.Base;

namespace Fusion.Connector.OpenHR.MessageComponents
{
	[System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "4.0.30319.17929")]
	[Serializable]
	[XmlType(AnonymousType = true, Namespace = "http://advancedcomputersoftware.com/xml/fusion/socialCare")]
	[XmlRoot(Namespace = "http://advancedcomputersoftware.com/xml/fusion/socialCare", IsNullable = false, ElementName = "staffPictureChange")]
	public class StaffPictureChange : BaseMessageComponent
	{

		public StaffPictureChange() { }

		public StaffPictureChange(Guid BusRef, Picture staffPicture)
		{
			staffRef = BusRef.ToString();
			data = new StaffPictureChangeData
			{
				pictureChange = staffPicture,
				recordStatus = staffPicture.isRecordInactive == true ? RecordStatusStandard.Inactive : RecordStatusStandard.Active,
				auditUserName = "OpenHR User"
			};
		}

		public StaffPictureChangeData data { get; set; }

		[XmlAttributeAttribute]
		public string staffRef { get; set; }
	}

}
