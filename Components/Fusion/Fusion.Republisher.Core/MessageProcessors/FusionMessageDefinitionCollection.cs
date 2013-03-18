using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace Fusion.Republisher.Core.MessageProcessors
{
    public class FusionMessageDefinitionCollection : IExtractFusionMessageData
    {
        public FusionMessageDefinition[] MessageData;
        public string XPath;

        public object ReadFromMessage(System.Xml.XmlNode message, System.Xml.XmlNamespaceManager ns, object currentState)
        {
            XmlNode messageRoot = message;
            if (XPath != null)
            {
                messageRoot = message.SelectSingleNode(XPath, ns);
            }

            if (messageRoot == null)
            {
                // Preserve current state if optional node not present

                return currentState ?? new FusionMessageDefinitionCollectionData
                {
                    Present = false,
                    Data = null
                };
            }

            FusionMessageDefinitionCollectionData messageData = currentState as FusionMessageDefinitionCollectionData;
            if (messageData == null)
            {
                messageData = new FusionMessageDefinitionCollectionData
                {
                    Present = true,
                    Data = new Dictionary<string, object>()
                };
            }

            messageData.Present = true;
            if (messageData.Data == null)
            {
                messageData.Data = new Dictionary<string, object>();
            }

            foreach (FusionMessageDefinition md in MessageData)
            {
                object contentData = null;
                messageData.Data.TryGetValue(md.Tag, out contentData);

                contentData = md.MessageExtractor.ReadFromMessage(messageRoot, ns, contentData);

                messageData.Data[md.Tag] = contentData;
            }

            return messageData;
        }

        public void Update(System.Xml.XmlNode message, System.Xml.XmlNamespaceManager ns, object currentState)
        {
            XmlNode messageRoot = message;
            if (XPath != null)
            {
                messageRoot = message.SelectSingleNode(XPath, ns);
            }

            if (messageRoot == null)
            {
                return;
            }

            FusionMessageDefinitionCollectionData messageData = currentState as FusionMessageDefinitionCollectionData;

            if (messageData == null || messageData.Present == false)
            {
                // Remove node from template

                messageRoot.ParentNode.RemoveChild(messageRoot);
                return;
            }

            foreach (FusionMessageDefinition md in MessageData)
            {
                object currentValue = null;
                messageData.Data.TryGetValue(md.Tag, out currentValue);

                md.MessageExtractor.Update(messageRoot, ns, currentValue);
            }
        }
    }

    public class FusionMessageDefinitionCollectionData
    {
        public bool Present;
        public Dictionary<string, object> Data;
    }
}
