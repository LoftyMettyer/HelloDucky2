﻿using System;
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

				[XmlElementAttribute(IsNullable = false)]
				public string documentReference { get; set; }

				[XmlElementAttribute(IsNullable = true)]
				public string requestedBy { get; set; }

				[XmlElementAttribute(DataType = "date", IsNullable = true)]
				public DateTime? requestedDate { get; set; }

				[XmlElementAttribute(IsNullable = true)]
				public string acceptedBy { get; set; }

				[XmlElementAttribute(DataType = "date", IsNullable = true)]
				public DateTime? acceptedDate { get; set; }

				[XmlIgnoreAttribute]
				public bool validFromSpecified { get; set; }

				[XmlIgnoreAttribute]
				public bool validToSpecified { get; set; }

				[XmlIgnoreAttribute]
				public bool requestedBySpecified { get; set; }

				[XmlIgnoreAttribute]
				public bool requestedDateSpecified { get; set; }

				[XmlIgnoreAttribute]
				public bool acceptedDateSpecified { get; set; }

				[XmlIgnoreAttribute]
				public bool acceptedBySpecified { get; set; }

				[XmlIgnoreAttribute]
				public int? id_Staff { get; set; }

				[XmlIgnoreAttribute]
				public bool? isRecordInactive { get; set; }

		}
}
