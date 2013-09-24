using System;
using Fusion.Connector.OpenHR.MessageComponents;
using Fusion.Core.Sql.OutboundBuilder;
using StructureMap.Attributes;
using Fusion.Core.Sql;
using Fusion.Connector.OpenHR.Configuration;
using Fusion.Messages.General;
using Fusion.Messages.SocialCare;
using Fusion.Connector.OpenHR.Database;

namespace Fusion.Connector.OpenHR.OutboundBuilders
{
    public class StaffPictureChangeMessageBuilder : IOutboundBuilder
    {

        [SetterProperty]
        public IBusRefTranslator refTranslator { get; set; }

				[SetterProperty]
				public IFusionConfiguration config { get; set; }

        private Type _myType;
        private string _messageType;

        public FusionMessage Build(SendFusionMessageRequest source)
        {
            var busRef = refTranslator.GetBusRef(EntityTranslationNames.Staff, source.LocalId);
            var picture = DatabaseAccess.readPicture(Convert.ToInt32(source.LocalId));

						var ChangeMessage = new StaffPictureChange(busRef, picture);

            _messageType = source.MessageType + "Request";
            _myType = Type.GetType("Fusion.Messages.SocialCare." + _messageType + ", Fusion.Messages.SocialCare");

            if (_myType != null)
            {
                var theMessage = (StaffPictureChangeRequest)Activator.CreateInstance(_myType);

                theMessage.Community = config.Community;
           
                theMessage.PrimaryEntityRef = busRef;
                theMessage.CreatedUtc = source.TriggerDate;
                theMessage.Id = Guid.NewGuid();
                theMessage.Originator = config.ServiceName;
                theMessage.EntityRef = busRef;
								theMessage.Xml = ChangeMessage.ToXml();

                return theMessage;
            }
            return null;
        }
    }
}
