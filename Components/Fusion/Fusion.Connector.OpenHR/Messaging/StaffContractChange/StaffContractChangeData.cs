using Fusion.Connector.OpenHR.MessageComponents.Enums;
using Fusion.Connector.OpenHR.MessageComponents.Component;
using System.Xml.Serialization;

namespace Fusion.Connector.OpenHR.MessageComponents.Data
{
    [System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "4.0.30319.17929")]
    [System.SerializableAttribute]
    [XmlTypeAttribute(AnonymousType = true, Namespace = "http://advancedcomputersoftware.com/xml/fusion/socialCare")]
    public class StaffContractChangeData
    {
        public StaffContractChangeData()
        {
            staffContract = new Contract();
        }

        public Contract staffContract { get; set; }

        [XmlAttributeAttribute]
        public string auditUserName { get; set; }

        [XmlAttributeAttribute]
        public RecordStatusStandard recordStatus { get; set; }
    }


}
