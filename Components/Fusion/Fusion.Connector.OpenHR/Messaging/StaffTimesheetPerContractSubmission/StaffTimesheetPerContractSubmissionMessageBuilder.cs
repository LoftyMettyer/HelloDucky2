using System;
using Fusion.Connector.OpenHR.MessageComponents.Data;
using Fusion.Core.Sql.OutboundBuilder;
using StructureMap.Attributes;
using Fusion.Core.Sql;
using Fusion.Connector.OpenHR.Configuration;
using Fusion.Messages.General;
using Fusion.Messages.SocialCare;
using Fusion.Connector.OpenHR.Database;
using Fusion.Connector.OpenHR.MessageComponents;
using Fusion.Connector.OpenHR.MessageComponents.Component;
using Fusion.Connector.OpenHR.MessageComponents.Enums;
using System.IO;

using System.Xml;
using System.Xml.Serialization;

namespace Fusion.Connector.OpenHR.OutboundBuilders
{
    public class StaffTimesheetPerContractSubmissionMessageBuilder : IOutboundBuilder
    {

        [SetterProperty]
        public IBusRefTranslator refTranslator { get; set; }

        [SetterProperty]
        public IFusionConfiguration config { get; set; }


        public FusionMessage Build(SendFusionMessageRequest source)
        {
            Guid timesheetRef = refTranslator.GetBusRef(EntityTranslationNames.Timesheet, source.LocalId);

            TimesheetPerContract timesheet = DatabaseAccess.readTimesheet(Convert.ToInt32(source.LocalId));

            var xsSubmit = new XmlSerializer(typeof(StaffTimesheetPerContractSubmission));
            var subReq = new StaffTimesheetPerContractSubmission();
            subReq.data = new StaffTimesheetPerContractSubmissionData();

            subReq.data.staffTimesheetPerContract = timesheet;
            subReq.data.recordStatus = RecordStatusTransactional.Active;

            subReq.data.auditUserName = "OpenHR user";

            Guid staffRef = refTranslator.GetBusRef(EntityTranslationNames.Staff, timesheet.id_Staff.ToString());

            subReq.submissionRef = timesheetRef.ToString();
            subReq.staffRef = staffRef.ToString();

            var sww = new StringWriter();
            XmlWriter writer = XmlWriter.Create(sww);
            xsSubmit.Serialize(writer, subReq);
            string xml = sww.ToString();

            string messageType = source.MessageType + "Request";
            Type myType = Type.GetType("Fusion.Messages.SocialCare." + messageType + ", Fusion.Messages.SocialCare");

            if (myType != null)
            {
                var theMessage = (StaffTimeSheetPerContractSubmissionMessage)Activator.CreateInstance(myType);

                theMessage.Community = config.Community;

                theMessage.PrimaryEntityRef = timesheetRef;
                theMessage.CreatedUtc = source.TriggerDate;
                theMessage.Id = Guid.NewGuid();
                theMessage.Originator = config.ServiceName;
                theMessage.EntityRef = staffRef;
                theMessage.Xml = xml;

                return theMessage;
            }
            return null;
        }
    }
}
