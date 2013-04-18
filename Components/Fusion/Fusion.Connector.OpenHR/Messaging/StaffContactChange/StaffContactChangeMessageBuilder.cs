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
using Fusion.Connector.OpenHR.MessageComponents.Enums;
using System.IO;

using System.Xml;
using System.Xml.Serialization;

namespace Fusion.Connector.OpenHR.OutboundBuilders
{
    public class StaffContactChangeMessageBuilder : IOutboundBuilder
    {
        [SetterProperty]
        public IBusRefTranslator refTranslator { get; set; }

        [SetterProperty]
        public IFusionConfiguration config { get; set; }

        public FusionMessage Build(SendFusionMessageRequest source)
        {
            var contactRef = refTranslator.GetBusRef(EntityTranslationNames.Contact, source.LocalId);

            var contact = DatabaseAccess.readContact(Convert.ToInt32(source.LocalId));

            var xsSubmit = new XmlSerializer(typeof(StaffContactChange));
            var subReq = new StaffContactChange();
            subReq.data = new StaffContactChangeData
                {
                    staffContact = contact,
                    recordStatus = contact.isRecordInactive == true ? RecordStatusStandard.Inactive : RecordStatusStandard.Active,
                    auditUserName = "OpenHR user"
                };

            var staffRef = refTranslator.GetBusRef(EntityTranslationNames.Staff, contact.id_Staff.ToString());

            subReq.staffContactRef = contactRef.ToString();
            subReq.staffRef = staffRef.ToString();

            var sww = new StringWriter();
            var writer = XmlWriter.Create(sww);
            xsSubmit.Serialize(writer, subReq);
            var xml = sww.ToString();

            var messageType = source.MessageType + "Request";
            var myType = Type.GetType("Fusion.Messages.SocialCare." + messageType + ", Fusion.Messages.SocialCare");

            if (myType != null)
            {
                var theMessage = (StaffContactChangeRequest)Activator.CreateInstance(myType);

                theMessage.Community = config.Community;

                theMessage.PrimaryEntityRef = staffRef ;
                theMessage.CreatedUtc = source.TriggerDate;
                theMessage.Id = Guid.NewGuid();
                theMessage.Originator = config.ServiceName;
                theMessage.EntityRef = contactRef;
                theMessage.Xml = xml;

                return theMessage;
            }

            return null;
        }
    }



}




