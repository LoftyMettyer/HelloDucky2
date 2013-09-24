using System;
using System.Configuration;
using Fusion.Connector.OpenHR.Configuration;
using Fusion.Connector.OpenHR.MessageComponents;
using Fusion.Core.Sql;
using Fusion.Core.Sql.OutboundBuilder;
using Fusion.Messages.General;
using Fusion.Messages.SocialCare;
using StructureMap.Attributes;
using Fusion.Connector.OpenHR.Database;

namespace Fusion.Connector.OpenHR.OutboundBuilders
{

    public class StaffChangeMessageBuilder : IOutboundBuilder
    {

			public string connectionString { get; set; }

			[SetterProperty]
			public IBusRefTranslator refTranslator { get; set; }

			[SetterProperty]
			public IFusionConfiguration config { get; set; }

        private Type _myType;
        private string _messageType;


        public FusionMessage Build(SendFusionMessageRequest source)
        {
            var busRef = refTranslator.GetBusRef(EntityTranslationNames.Staff, source.LocalId);
            var staff = DatabaseAccess.readStaff(Convert.ToInt32(source.LocalId));

						var ChangeMessage = new StaffChange(busRef, staff);

            _messageType = source.MessageType + "Request";
            _myType = Type.GetType("Fusion.Messages.SocialCare." + _messageType + ", Fusion.Messages.SocialCare");

            var theMessage = (StaffChangeRequest)Activator.CreateInstance(_myType);

            theMessage.Community = ConfigurationManager.AppSettings["Community"];

            theMessage.PrimaryEntityRef = busRef;
            theMessage.CreatedUtc = source.TriggerDate;
            theMessage.Id = Guid.NewGuid();
            theMessage.Originator = config.ServiceName;
            theMessage.EntityRef = busRef;
						theMessage.Xml = ChangeMessage.ToXml();

            return theMessage;

        }

    
    }
}
