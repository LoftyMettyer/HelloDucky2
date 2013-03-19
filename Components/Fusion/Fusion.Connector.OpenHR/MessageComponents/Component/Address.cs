
namespace Fusion.Connector.OpenHR.MessageComponents.Component
{

    [System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "4.0.30319.17929")]
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(AnonymousType = true, Namespace = "http://advancedcomputersoftware.com/xml/fusion/socialCare")]
    public partial class Address
    {

        [System.Xml.Serialization.XmlElementAttribute(IsNullable = false)]
        public string addressLine1 { get; set; }

        [System.Xml.Serialization.XmlElementAttribute(IsNullable = false)]
        public string addressLine2 { get; set; }
  
        [System.Xml.Serialization.XmlElementAttribute(IsNullable = false)]
        public string addressLine3  { get; set; }

        [System.Xml.Serialization.XmlElementAttribute(IsNullable = false)]
        public string addressLine4 { get; set; }

        [System.Xml.Serialization.XmlElementAttribute(IsNullable = false)]
        public string addressLine5  { get; set; }

        [System.Xml.Serialization.XmlElementAttribute(IsNullable = false)]
        public string postCode  { get; set; }
    }
}
