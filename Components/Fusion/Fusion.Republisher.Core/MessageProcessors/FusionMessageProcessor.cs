using Fusion.Republisher.Core.MessageStateSerializer;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;

namespace Fusion.Republisher.Core.MessageProcessors
{
    public class FusionMessageProcessor : Fusion.Republisher.Core.MessageProcessors.IFusionMessageProcessor
    {
        public void UpdateStateFromMessage(IMessageDefinition messageDefinition, MessagePersistedState currentState, string messageXml)
        {
            XmlDocument doc = new XmlDocument();
            doc.LoadXml(messageXml);

            // Scan through definition and update with data from message

            foreach (FusionMessageDefinition md in messageDefinition.MessageDefinition)
            {
                object contentData = null;
                currentState.State.TryGetValue(md.Tag, out contentData);

                contentData = md.MessageExtractor.ReadFromMessage(doc, messageDefinition.NamespaceManager, contentData);

                currentState.State[md.Tag] = contentData;
            }
        }

        public string CreateMessageFromState(IMessageDefinition messageDefinition, MessagePersistedState currentState)
        {
            // Generate message

            XmlDocument template = new XmlDocument();
            template.Load(messageDefinition.GetBlankXmlTemplate());

            foreach (FusionMessageDefinition md in messageDefinition.MessageDefinition)
            {
                object currentValue = null;

                currentState.State.TryGetValue(md.Tag, out currentValue);

                md.MessageExtractor.Update(template, messageDefinition.NamespaceManager, currentValue);
            }

            StringWriter outXml = new StringWriter();
            template.Save(outXml);

            return outXml.ToString();
        }
    }
}
