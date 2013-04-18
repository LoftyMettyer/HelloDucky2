using System;
using System.Xml.Serialization;
using Fusion.Connector.OpenHR.MessageComponents.Component;
using Fusion.Connector.OpenHR.MessageComponents.Enums;

namespace Fusion.Connector.OpenHR.MessageComponents.Data
{
    [SerializableAttribute]
    [XmlTypeAttribute(AnonymousType = true, Namespace = "http://advancedcomputersoftware.com/xml/fusion/socialCare")]
    public class StaffContactChangeData
    {
        public StaffContactChangeData()
        {
            staffContact = new Contact();
        }

        public Contact staffContact { get; set; }

        [XmlAttributeAttribute]
        public string auditUserName { get; set; }

        [XmlAttributeAttribute]
        public RecordStatusStandard recordStatus { get; set; }
    }


}
