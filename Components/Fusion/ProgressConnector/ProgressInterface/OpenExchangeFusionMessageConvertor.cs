using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Fusion.Messages.General;
using System.Xml;
using Fusion.Core;

namespace Connector1.ProgressInterface
{
    public class OpenExchangeFusionMessageConvertor : Connector1.ProgressInterface.IOpenExchangeFusionMessageConvertor
    {
        public string BuildOpenExchangeMessage(FusionMessage message)
        {
            StringBuilder s = new StringBuilder();

            XmlWriterSettings writerSettings = new XmlWriterSettings();
            writerSettings.OmitXmlDeclaration = true;

            XmlWriter x = XmlWriter.Create(s, writerSettings);

            x.WriteStartElement("fusionMessage", "http://advancedcomputersoftware.com/xml/fusion");

            x.WriteElementString("messageType", message.GetMessageName());
            x.WriteElementString("id", message.Id.ToString());
            x.WriteElementString("originator", message.Originator);
            x.WriteStartElement("payload");

            x.WriteRaw(message.Xml);

            x.WriteEndElement(); // payload
            
            x.WriteElementString("created", XmlConvert.ToString(message.CreatedUtc, XmlDateTimeSerializationMode.Utc));
            x.WriteElementString("schemaVersion", message.SchemaVersion.ToString());
            if (message.EntityRef.HasValue) {
                x.WriteElementString("entityRef", message.EntityRef.Value.ToString());
            }
            
            x.Close();

            return s.ToString();
        }
    }
}
