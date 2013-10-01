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
    public class StaffContractChangeMessageBuilder : IOutboundBuilder
    {
        [SetterProperty]
        public IBusRefTranslator refTranslator { get; set; }

        [SetterProperty]
        public IFusionConfiguration config { get; set; }

        public FusionMessage Build(SendFusionMessageRequest source)
        {
            var contractRef = refTranslator.GetBusRef(EntityTranslationNames.Contract, source.LocalId);
            var contract = DatabaseAccess.readContract(Convert.ToInt32(source.LocalId));
						var staffRef = refTranslator.GetBusRef(EntityTranslationNames.Staff, contract.id_Staff.ToString());

						var ChangeMessage = new StaffContractChange(contractRef, staffRef, contract);

            string messageType = source.MessageType + "Request";
            Type myType = Type.GetType("Fusion.Messages.SocialCare." + messageType + ", Fusion.Messages.SocialCare");

            if (myType != null)
            {
                var theMessage = (StaffContractChangeRequest)Activator.CreateInstance(myType);

                theMessage.Community = config.Community;
                theMessage.PrimaryEntityRef = staffRef;
                theMessage.CreatedUtc = source.TriggerDate;
                theMessage.Id = Guid.NewGuid();
                theMessage.Originator = config.ServiceName;
                theMessage.EntityRef = contractRef;
								theMessage.Xml = ChangeMessage.ToXml();

                return theMessage;
            }
            return null;
        }
    }
}
