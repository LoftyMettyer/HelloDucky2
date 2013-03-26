
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Fusion.Core.Sql;
using StructureMap.Attributes;
using Fusion.Core.OutboundFilters;
using Fusion.Messages.General;
using Fusion.Messages.SocialCare;
using Fusion.Core.MessageValidators;
using Fusion.Core;
using log4net;


namespace Fusion.Connector.OpenHR.Messaging
{
    public abstract class SchemaValidatorOutboundFilterHandler<T> : OutboundFilterHandler<T> where T : FusionMessage
    {
        protected ILog Logger;

        [SetterProperty]
        public IMessageTracking MessageTracking
        {
            get;
            set;
        }

        public SchemaValidatorOutboundFilterHandler()
        {
            Logger = LogManager.GetLogger(this.GetType());
        }


        protected bool checkAlreadySent(T message)
        {
            string messageType = message.GetMessageName();

            if (message.EntityRef != null)
            {
                var busRef = message.EntityRef.Value;

                string lastMessageGenerated = MessageTracking.GetLastGeneratedXml(messageType, busRef);

                // Simple compare, could be more elaborate
                if (lastMessageGenerated == message.Xml)
                {
                    Logger.InfoFormat("Outbound Message {0} message previously sent {1} ", messageType, busRef);
                    return true;
                }
            }

            return false;
        }


        protected bool CheckValidity(string xml, string schemaName)
        {
            SchemaValidator sv = new SchemaValidator(
                new EmbeddedXmlResourceResolver(),
                "http://advancedcomputersoftware.com/xml/fusion/socialCare",
                "res://Fusion.Connector.OpenHR/Fusion.Connector.OpenHR.Schemas/");

            var validation = sv.Validate(xml, schemaName);

            if (validation.HasErrors)
            {
                Logger.InfoFormat("Outbound Message {0} validity failed - Error = {1}", schemaName, validation.ValidationErrorString);
            }

            return !validation.HasErrors;
        }

        protected bool CheckValidity(T message)
        {
            string schemaName = message.GetMessageName();

            schemaName = schemaName + ".xsd";
            schemaName = Char.ToLower(schemaName[0]) + schemaName.Substring(1);

            return CheckValidity(message.Xml, schemaName);
        }

        public override bool Handle(T message)
        {
            bool messageValid = CheckValidity(message);

            Logger.InfoFormat("Outbound Message {0} schema check status {1}", message.GetType().Name, messageValid);

            return messageValid;
        }

    }
}



//namespace Fusion.Connector.OpenHR.Messaging
//{
//    public class ServiceUserUpdateMessageSchemaValidatorOutboundFilter : Fusion.Connector.OpenHR.Messaging.SchemaValidatorOutboundFilterHandler<StaffChangeRequest>
//    {
//        public override bool Handle(StaffChangeRequest message)
//        {
//            bool valid = base.CheckValidity(message);
//            return valid;
//        }
//    }
//}