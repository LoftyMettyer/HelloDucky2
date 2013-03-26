using System.Xml.Serialization;
using Fusion.Connector.OpenHR.MessageComponents.Enums;
using Fusion.Connector.OpenHR.MessageComponents.Component;

namespace Fusion.Connector.OpenHR.MessageComponents.Data
{
    [System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "4.0.30319.17929")]
    [System.SerializableAttribute]
    [XmlTypeAttribute(AnonymousType = true, Namespace = "http://advancedcomputersoftware.com/xml/fusion/socialCare")]
    public class StaffSkillChangeData
    {
        public StaffSkillChangeData()
        {
            staffSkill = new Skill();
        }

        public Skill staffSkill { get; set; }

        [XmlAttributeAttribute]
        public string auditUserName { get; set; }

        [XmlAttributeAttribute]
        public RecordStatusRescindable recordStatus { get; set; }

    }

}
