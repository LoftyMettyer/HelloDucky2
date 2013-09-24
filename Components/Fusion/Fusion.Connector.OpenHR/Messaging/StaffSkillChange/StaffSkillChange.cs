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
	[XmlRoot(Namespace = "http://advancedcomputersoftware.com/xml/fusion/socialCare", IsNullable = false, ElementName = "staffSkillChange")]
	public class StaffSkillChange : BaseMessageComponent
	{

		public StaffSkillChange() { }

		public StaffSkillChange(Guid busRef, Guid? parentRef, Skill skill)
		{
			StaffSkillRef = busRef.ToString();
			staffRef = parentRef.ToString();
			data = new StaffSkillChangeData
			{
				staffSkill = skill,
				recordStatus = skill.isRecordInactive == true ? RecordStatusStandard.Inactive : RecordStatusStandard.Active,
				auditUserName = "OpenHR User"
			};
		}

		public StaffSkillChangeData data { get; set; }

		[XmlAttributeAttribute(AttributeName = "staffSkillRef")]
		public string StaffSkillRef { get; set; }

		[XmlAttributeAttribute]
		public string staffRef { get; set; }
	}


}
