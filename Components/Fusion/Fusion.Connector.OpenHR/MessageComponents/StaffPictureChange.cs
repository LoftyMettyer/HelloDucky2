using System;
using Fusion.Connector.OpenHR.MessageComponents.Data;
using System.Xml.Serialization;

namespace Fusion.Connector.OpenHR.MessageComponents
{
    [System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "4.0.30319.17929")]
    [Serializable]
    [XmlType(AnonymousType = true, Namespace = "http://advancedcomputersoftware.com/xml/fusion/socialCare")]
    [XmlRoot(Namespace = "http://advancedcomputersoftware.com/xml/fusion/socialCare", IsNullable = false, ElementName = "staffPictureChange")]
    public class StaffPictureChange
    {

        public StaffPictureChangeData data { get; set; }

        [XmlAttributeAttribute]
        public int version { get; set; }

        [XmlAttributeAttribute]
        public string staffRef { get; set; }
    }

}
