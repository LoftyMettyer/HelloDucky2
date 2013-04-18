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

            var xsSubmit = new XmlSerializer(typeof(StaffContractChange));
            var subReq = new StaffContractChange();
            subReq.data = new StaffContractChangeData
                {
                    staffContract = contract,
                    recordStatus = contract.isRecordInactive == true ? RecordStatusStandard.Inactive: RecordStatusStandard.Active,
                    auditUserName = "OpenHR user"
                };

            var staffRef = refTranslator.GetBusRef(EntityTranslationNames.Staff, contract.id_Staff.ToString());

            subReq.staffContractRef = contractRef.ToString();
            subReq.staffRef = staffRef.ToString();

            var sww = new StringWriter();
            XmlWriter writer = XmlWriter.Create(sww);
            xsSubmit.Serialize(writer, subReq);
            string xml = sww.ToString();

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
                theMessage.Xml = xml;

                return theMessage;
            }
            return null;
        }
    }
}
