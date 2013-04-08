using System;
using System.Xml.Serialization;
using Fusion.Connector.OpenHR.MessageComponents.Enums;

namespace Fusion.Connector.OpenHR.MessageComponents.Component
{
    [System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "4.0.30319.17929")]
    [SerializableAttribute]
    [XmlTypeAttribute(AnonymousType = true, Namespace = "http://advancedcomputersoftware.com/xml/fusion/socialCare")]
    public class LegalDocument
    {

        public LegalDocumentTypes? typeName { get; set; }

        [XmlElementAttribute(DataType = "date", IsNullable = true)]
        public DateTime? validFrom { get; set; }

        [XmlElementAttribute(DataType = "date", IsNullable = true)]
        public DateTime? validTo { get; set; }

        public string documentReference { get; set; }

        [XmlElementAttribute(IsNullable = true)]
        public string secondaryReference { get; set; }

        [XmlElementAttribute(IsNullable = true)]
        public string requestedBy { get; set; }

        [XmlElementAttribute(DataType = "date", IsNullable = true)]
        public DateTime? requestedDate { get; set; }

        [XmlElementAttribute(IsNullable = true)]
        public string acceptedBy { get; set; }

        [XmlElementAttribute(DataType = "date", IsNullable = true)]
        public DateTime? acceptedDate { get; set; }

        [XmlIgnoreAttribute]
        public bool acceptedDateFieldSpecified { get; set; }

        [XmlIgnoreAttribute]
        public bool requestedDateFieldSpecified { get; set; }


        [XmlIgnoreAttribute]
        public int? id_Staff { get; set; }

    }
}
