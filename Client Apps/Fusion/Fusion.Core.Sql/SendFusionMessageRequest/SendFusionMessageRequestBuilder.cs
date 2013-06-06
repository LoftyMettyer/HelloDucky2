using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Xml.Linq;
using System.Xml;

namespace Fusion.Core.Sql
{
    public class SendFusionMessageRequestBuilder : ISendFusionMessageRequestBuilder
    {
        public SendFusionMessageRequest Build(Stream bodyStream)
        {
            XDocument loaded = XDocument.Load(bodyStream);

            var rootNode = loaded.Element("SendFusionMessage");
            string msgType = (string)rootNode.Element("MessageType");
            string localId = (string)rootNode.Element("LocalId");

            DateTime dt =
                XmlConvert.ToDateTime(
                    (string)rootNode.Element("TriggerDate"),
                    XmlDateTimeSerializationMode.Utc
                 );

            return new SendFusionMessageRequest {
                MessageType = msgType,
                LocalId = localId,
                TriggerDate = dt,
            };                               
        }
    }
}
