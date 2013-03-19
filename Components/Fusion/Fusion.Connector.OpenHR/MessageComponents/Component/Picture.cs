using Fusion.Connector.OpenHR.MessageComponents.Enums;
using System.Xml.Serialization;

namespace Fusion.Connector.OpenHR.MessageComponents.Component
{
    [System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "4.0.30319.17929")]
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [XmlTypeAttribute(AnonymousType = true, Namespace = "http://advancedcomputersoftware.com/xml/fusion/socialCare")]
    public class Picture
    {
        public ImageType imageType { get; set; }

        [XmlElementAttribute(IsNullable = false)]
        public byte[] picture {get; set; }
    }

}
