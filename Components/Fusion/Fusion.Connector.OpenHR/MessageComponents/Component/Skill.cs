using System;

namespace Fusion.Connector.OpenHR.MessageComponents.Component
{
    public class Skill
    {
        public string name { get; set; }

        public DateTime trainingStart { get; set; }

        [System.Xml.Serialization.XmlIgnoreAttribute]
        public bool trainingStartSpecified { get; set; }

        public DateTime? trainingEnd { get; set; }

        [System.Xml.Serialization.XmlIgnoreAttribute]
        public bool trainingEndSpecified { get; set; }

        public DateTime validFrom { get; set; }

        [System.Xml.Serialization.XmlIgnoreAttribute]
        public bool validFromSpecified { get; set; }

        public DateTime? validTo { get; set; }

        [System.Xml.Serialization.XmlIgnoreAttribute]
        public bool validToSpecified { get; set; }

        public string reference { get; set; }

        public string outcome { get; set; }

        public bool? didNotAttend { get; set; }

        [System.Xml.Serialization.XmlIgnoreAttribute]
        public bool didNotAttendSpecified { get; set; }
    }


}
