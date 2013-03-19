﻿using System;
using System.Xml.Serialization;
using Fusion.Connector.OpenHR.MessageComponents.Data;

namespace Fusion.Connector.OpenHR.MessageComponents
{
    [System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "4.0.30319.17929")]
    [SerializableAttribute]
    [XmlType(AnonymousType = true, Namespace = "http://advancedcomputersoftware.com/xml/fusion/socialCare")]
    [XmlRoot(Namespace = "http://advancedcomputersoftware.com/xml/fusion/socialCare", IsNullable = false, ElementName = "staffTimesheetPerContractSubmission")]
    public class StaffTimesheetPerContractSubmission
    {

        public StaffTimesheetPerContractSubmissionData data { get; set; }

        [XmlAttributeAttribute]
        public int version { get; set; }

        [XmlAttributeAttribute]
        public string submissionRef { get; set; }

        [XmlAttributeAttribute]
        public string staffRef { get; set; }
    }


}
