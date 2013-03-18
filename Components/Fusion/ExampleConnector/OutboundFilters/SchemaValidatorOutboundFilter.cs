using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Fusion.Core.Sql;
using StructureMap.Attributes;
using Fusion.Core.OutboundFilters;
using Fusion.Messages.General;
using Fusion.Core.MessageValidators;
using Fusion.Core;
using log4net;

namespace Connector1.OutboundFilters
{
    public abstract class SchemaValidatorOutboundFilterHandler<T> : OutboundFilterHandler<T> where T : FusionMessage
    {
        protected ILog Logger;

        public SchemaValidatorOutboundFilterHandler()
        {
            Logger = LogManager.GetLogger(this.GetType());
        }

        protected bool CheckValidity(string xml, string schemaName)
        {
            SchemaValidator sv = new SchemaValidator(
                new EmbeddedXmlResourceResolver(), 
                "http://advancedcomputersoftware.com/xml/fusion", 
                "res://ExampleConnector/ExampleConnector/Schemas/");

            var validation = sv.Validate(xml, schemaName);

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
