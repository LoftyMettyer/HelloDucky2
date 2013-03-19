using System;
using Fusion.Connector.OpenHR.MessageComponents.Enums;
using Fusion.Connector.OpenHR.MessageComponents.Component;
using System.Xml.Serialization;

namespace Fusion.Connector.OpenHR.MessageComponents.Data
{
    [System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "4.0.30319.17929")]
    [SerializableAttribute]
    [XmlTypeAttribute(AnonymousType = true, Namespace = "http://advancedcomputersoftware.com/xml/fusion/socialCare")]
    public class StaffLegalDocumentChangeData
    {
        public StaffLegalDocumentChangeData()
        {
            staffLegalDocument = new LegalDocument();
        }

        public LegalDocument staffLegalDocument { get; set; }

        [XmlAttributeAttribute]
        public string auditUserName { get; set; }

        [XmlAttributeAttribute]
        public RecordStatusRescindable recordStatus { get; set; }
    }


}
