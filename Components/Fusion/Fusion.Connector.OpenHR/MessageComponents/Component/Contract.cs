using System;
using System.Xml.Serialization;

namespace Fusion.Connector.OpenHR.MessageComponents.Component
{
    public partial class Contract
    {
        public string contractName { get; set; }

        [XmlElement(IsNullable = true)]
        public string department { get; set; }

        public string primarySite { get; set; }

        public decimal? contractedHoursPerWeek { get; set; }

        public decimal? maximumHoursPerWeek { get; set; }

        [XmlElementAttribute(DataType = "date")]
        public DateTime? effectiveFrom { get; set; }

        [XmlIgnoreAttribute()]
        public bool effectiveFromSpecified { get; set; }

        [XmlElementAttribute(DataType = "date")]
        public DateTime? effectiveTo { get; set; }

        [XmlIgnoreAttribute()]
        public bool effectiveToSpecified { get; set; }

        [XmlIgnoreAttribute]
        public int? id_Staff { get; set; }

    }

}
