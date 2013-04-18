using System.Xml.Serialization;
using Fusion.Connector.OpenHR.MessageComponents.Component;
using Fusion.Connector.OpenHR.MessageComponents.Enums;

namespace Fusion.Connector.OpenHR.MessageComponents.Data
{
    public class StaffTimesheetPerContractSubmissionData
    {

        public TimesheetPerContract staffTimesheetPerContract { get; set; }

        [XmlAttributeAttribute]
        public string auditUserName { get; set; }

        [XmlAttributeAttribute]
        public RecordStatusTransactional recordStatus { get; set; }
    }

}
