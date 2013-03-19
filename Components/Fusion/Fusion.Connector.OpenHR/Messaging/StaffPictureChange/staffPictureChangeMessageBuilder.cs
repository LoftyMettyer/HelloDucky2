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
using Fusion.Connector.OpenHR.MessageComponents.Data;
using Fusion.Connector.OpenHR.MessageComponents.Component;
using Fusion.Connector.OpenHR.MessageComponents.Enums;
using System.IO;

using System.Xml;
using System.Xml.Serialization;

namespace Fusion.Connector.OpenHR.Messaging.StaffPictureChange
{
    public class StaffPictureChangeMessageBuilder : IOutboundBuilder
    {

        public string connectionString { get; set; }

        [SetterProperty]
        public IBusRefTranslator refTranslator { get; set; }

        [SetterProperty]
        public IFusionConfiguration config { get; set; }

        private Type _myType;
        private string _messageType;


        public FusionMessage Build(Core.Sql.SendFusionMessageRequest source)
        {
            string xml = "";

            Guid busRef = refTranslator.GetBusRef(EntityTranslationNames.Staff, source.LocalId);
            Picture picture = DatabaseAccess.readPicture(Convert.ToInt32(source.LocalId));

            var xsSubmit = new XmlSerializer(typeof(MessageComponents.StaffPictureChange));
            var subReq = new MessageComponents.StaffPictureChange();

            subReq.data = new StaffPictureChangeData();

            subReq.data.pictureChange = picture;

            subReq.data.recordStatus = RecordStatusRescindable.Active;
            subReq.data.auditUserName = "OpenHR user";
            subReq.staffRef = busRef.ToString();

            var sww = new StringWriter();
            XmlWriter writer = XmlWriter.Create(sww);
            xsSubmit.Serialize(writer, subReq);
            xml = sww.ToString();

            _messageType = source.MessageType + "Request";
            _myType = Type.GetType("Fusion.Messages.SocialCare." + _messageType + ", Fusion.Messages.SocialCare");

            var theMessage = (StaffPictureChangeRequest)Activator.CreateInstance(_myType);

            theMessage.Community = config.Community;
           
            theMessage.PrimaryEntityRef = busRef;
            theMessage.CreatedUtc = source.TriggerDate;
            theMessage.Id = Guid.NewGuid();
            theMessage.Originator = config.ServiceName;
            theMessage.EntityRef = busRef;
            theMessage.Xml = xml;

            return theMessage;

        }
    }
}
