using System;
using System.Xml.Serialization;

namespace Fusion.Connector.OpenHR.MessageComponents.Component
{
    public class Skill
    {
        public string name { get; set; }

        [XmlElementAttribute(DataType = "date")]
        public DateTime? trainingStart { get; set; }

        [XmlIgnoreAttribute]
        public bool trainingStartSpecified { get; set; }

        [XmlElementAttribute(DataType = "date")]
        public DateTime? trainingEnd { get; set; }

        [XmlIgnoreAttribute]
        public bool trainingEndSpecified { get; set; }

        [XmlElementAttribute(DataType = "date")]
        public DateTime? validFrom { get; set; }

        [XmlIgnoreAttribute]
        public bool validFromSpecified { get; set; }

        [XmlElementAttribute(DataType = "date")]
        public DateTime? validTo { get; set; }

        [XmlIgnoreAttribute]
        public bool validToSpecified { get; set; }

        public string reference { get; set; }

        public string outcome { get; set; }

        public bool? didNotAttend { get; set; }

        [XmlIgnoreAttribute]
        public bool didNotAttendSpecified { get; set; }

        [XmlIgnoreAttribute]
        public int? id_Staff { get; set; }

    }


}
