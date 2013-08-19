﻿using System.Xml.Serialization;

namespace Fusion.Connector.OpenHR.MessageComponents.Component
{
    public class Contact
    {

        public string title { get; set; }

        [XmlElement(IsNullable = true)]
        public string forenames { get; set; }

        [XmlElement(IsNullable = true)]
        public string surname { get; set; }

        [XmlElement(IsNullable = true)]
        public string contactType { get; set; }

        [XmlElementAttribute(IsNullable = true)]
        public string relationshipType { get; set; }

        [XmlElement(IsNullable = true)]
        public string workMobile { get; set; }

        [XmlElement(IsNullable = true)]
        public string personalMobile { get; set; }

        [XmlElement(IsNullable = true)]
        public string workPhoneNumber { get; set; }

        [XmlElement(IsNullable = true)]
        public string homePhoneNumber { get; set; }

        [XmlElement(IsNullable = true)]
        public string email { get; set; }

        [XmlElement(IsNullable = true)]
        public string notes { get; set; }

        [XmlIgnoreAttribute]
        public bool homeAddressSpecified { get; set; }

        [XmlElementAttribute(IsNullable = true)]
        public Address homeAddress { get; set; }

        [XmlIgnoreAttribute]
        public int? id_Staff { get; set; }

        [XmlIgnoreAttribute]
        public bool? isRecordInactive { get; set; }


    }
}
