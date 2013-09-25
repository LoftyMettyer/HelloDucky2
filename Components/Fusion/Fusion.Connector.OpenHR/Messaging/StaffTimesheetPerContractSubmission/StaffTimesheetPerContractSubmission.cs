using System;
using System.CodeDom.Compiler;
using System.Xml.Serialization;
using Fusion.Connector.OpenHR.MessageComponents.Data;
using Fusion.Connector.OpenHR.Messaging.Base;

namespace Fusion.Connector.OpenHR.MessageComponents
{
    [GeneratedCode("xsd", "4.0.30319.17929")]
    [Serializable]
    [XmlType(AnonymousType = true, Namespace = "http://advancedcomputersoftware.com/xml/fusion/socialCare")]
		[XmlRoot(Namespace = "http://advancedcomputersoftware.com/xml/fusion/socialCare", IsNullable = false, ElementName = "staffTimesheetPerContractSubmission")]

	public class StaffTimesheetPerContractSubmission : BaseMessageComponent
    {

        public StaffTimesheetPerContractSubmissionData data { get; set; }

       	[XmlAttributeAttribute(AttributeName = "submissionRef")]
        public string SubmissionRef { get; set; }

        [XmlAttributeAttribute]
        public string staffRef { get; set; }
    }
}

