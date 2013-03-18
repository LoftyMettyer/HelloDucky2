using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using System.IO;
using System.Xml;

namespace ProgressConnector.ProgressInterface
{
    public class OpenExchangeMessageDecoder : IOpenExchangeMessageDecoder
    {
        public OpenExchangeGeneratedContent Decode(RawOpenExchangeData data)
        {
            XDocument loaded = XDocument.Load(new StringReader(data.MessageRequestXml));

            switch (loaded.Root.Name.LocalName)
            {
                case "fusionMessage":
                    return ParseMessage(loaded);
                case "fusionIdTranslation":
                    return ParseIdTranslation(loaded);
                case "fusionLog":
                    return ParseLog(loaded);

            }

            throw new ArgumentException("Invalid xml passed for conversion", "data");
        }

        private  OpenExchangeIdTranslation ParseIdTranslation(XDocument loaded)
        {
            XNamespace ns = "http://advancedcomputersoftware.com/xml/fusion";

            var rootNode = loaded.Element(ns + "fusionIdTranslation");

            string id = (string)rootNode.Element(ns + "id");
            string source = (string)rootNode.Element(ns + "source");
            string messageId = (string)rootNode.Element(ns + "messageId");
            string timeUtc = (string)rootNode.Element(ns + "timeUtc");
            string entityName = (string)rootNode.Element(ns + "entityName");
            string busRef = (string)rootNode.Element(ns + "busRef");
            string localId = (string)rootNode.Element(ns + "localId");
            string community = (string)rootNode.Element(ns + "community");      

            return new OpenExchangeIdTranslation
            {
                Id = new Guid(id),
                BusRef = new Guid(busRef),
                MessageId = String.IsNullOrEmpty(messageId) ? null : (Guid?)new Guid(messageId),
                TimeUtc = XmlConvert.ToDateTime(timeUtc, XmlDateTimeSerializationMode.Utc),
                EntityName = entityName,
                LocalId = localId,
                Community = community,
                Source = source
            };
        }

        private OpenExchangeLogMessage ParseLog(XDocument loaded)
        {
            XNamespace ns = "http://advancedcomputersoftware.com/xml/fusion";

            var rootNode = loaded.Element(ns + "fusionLog");

            string id = (string)rootNode.Element(ns + "id");
            string source = (string)rootNode.Element(ns + "source");
            string messageId = (string)rootNode.Element(ns + "messageId");
            string entityRef = (string)rootNode.Element(ns + "entityRef");
            string primaryEntityRef = (string)rootNode.Element(ns + "primaryEntityRef");
            string timeUtc = (string)rootNode.Element(ns + "timeUtc");
            string logLevel = (string)rootNode.Element(ns + "logLevel");
            string message = (string)rootNode.Element(ns + "message");
            string messageDescription = (string)rootNode.Element(ns + "messageDescription");
            string community = (string)rootNode.Element(ns + "community");

            return new OpenExchangeLogMessage
            {
                Id = new Guid(id),
                Source = source,
                
                MessageId = String.IsNullOrEmpty(messageId) ? null : (Guid?)new Guid(messageId),

                EntityRef = String.IsNullOrEmpty(entityRef) ? null : (Guid?)new Guid(entityRef),
                PrimaryEntityRef = String.IsNullOrEmpty(entityRef) ? null : (Guid?)new Guid(primaryEntityRef),

                TimeUtc = XmlConvert.ToDateTime(timeUtc, XmlDateTimeSerializationMode.Utc),

                LogLevel = logLevel,
                Message = message,
                MessageDescription = messageDescription,

                Community = community
            };
        }

        private  OpenExchangeMessage ParseMessage(XDocument loaded)
        {
            XNamespace ns = "http://advancedcomputersoftware.com/xml/fusion";

            var rootNode = loaded.Element(ns + "fusionMessage");

            string messageType = (string)rootNode.Element(ns + "messageType");
            string id = (string)rootNode.Element(ns + "id");
            string originator = (string)rootNode.Element(ns + "originator");

            string created = (string)rootNode.Element(ns + "created");
            string schemaVersion = (string)rootNode.Element(ns + "schemaVersion");
            string entityRef = (string)rootNode.Element(ns + "entityRef");
            string primaryEntityRef = (string)rootNode.Element(ns + "primaryEntityRef");
            string community = (string)rootNode.Element(ns + "community");      

            var payload = rootNode.Element(ns + "payload").CreateReader();
            payload.MoveToContent();

            return new OpenExchangeMessage
            {
                MessageType = messageType,
                Id = new Guid(id),
                Originator = originator,
                Xml = payload.ReadInnerXml(),
                Created = XmlConvert.ToDateTime(created, XmlDateTimeSerializationMode.Utc),
                SchemaVersion = Convert.ToInt32(schemaVersion),
                EntityRef = String.IsNullOrEmpty(entityRef) ? null : (Guid?)new Guid(entityRef),
                PrimaryEntityRef = String.IsNullOrEmpty(primaryEntityRef) ? null : (Guid?)new Guid(primaryEntityRef),

                Community = community
            };
        }
    }
}
