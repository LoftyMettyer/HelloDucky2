using System;
using Fusion.Core.Sql.OutboundBuilder;
using StructureMap.Attributes;
using Fusion.Core.Sql;
using Fusion.Connector.OpenHR.Configuration;
using Fusion.Messages.General;
using Fusion.Messages.SocialCare;
using Fusion.Connector.OpenHR.Database;
using Fusion.Connector.OpenHR.MessageComponents;

namespace Fusion.Connector.OpenHR.OutboundBuilders
{
    public class StaffLegalDocumentChangeMessageBuilder : IOutboundBuilder
    {
        [SetterProperty]
        public IBusRefTranslator refTranslator { get; set; }

        [SetterProperty]
        public IFusionConfiguration config { get; set; }

        public FusionMessage Build(SendFusionMessageRequest source)
        {
            var docRef = refTranslator.GetBusRef(EntityTranslationNames.Document, source.LocalId);
            var doc = DatabaseAccess.readDocument(Convert.ToInt32(source.LocalId));
            var staffRef = refTranslator.GetBusRef(EntityTranslationNames.Staff, doc.id_Staff.ToString());
						var ChangeMessage = new StaffLegalDocumentChange(docRef, staffRef, doc);
            var messageType = source.MessageType + "Request";
            var myType = Type.GetType("Fusion.Messages.SocialCare." + messageType + ", Fusion.Messages.SocialCare");

            var theMessage = (StaffLegalDocumentChangeRequest)Activator.CreateInstance(myType);

            theMessage.Community = config.Community;

            theMessage.PrimaryEntityRef = staffRef ;
            theMessage.CreatedUtc = source.TriggerDate;
            theMessage.Id = Guid.NewGuid();
            theMessage.Originator = config.ServiceName;
            theMessage.EntityRef = docRef;
						theMessage.Xml = ChangeMessage.ToXml();

            return theMessage;
        }
    }
}



