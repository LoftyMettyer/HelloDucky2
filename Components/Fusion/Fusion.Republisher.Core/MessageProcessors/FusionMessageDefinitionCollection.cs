using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using log4net;

namespace Fusion.Republisher.Core.MessageProcessors
{
    public class FusionMessageDefinitionCollection : IExtractFusionMessageData
    {
        const string XSISchema = "http://www.w3.org/2001/XMLSchema-instance";
        const string NilAttribute = "nil";

        public FusionMessageDefinition[] MessageData;
        public string XPath;
        public InitialState InitialState;

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

            if (messageRoot.Attributes[NilAttribute, XSISchema] != null)
            {
                messageData.Present = false;
                messageData.Data = new Dictionary<string, object>();
            }
            else
            {
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

            XmlDocument doc = message as XmlDocument;
            if (doc == null)
            {
                doc = message.OwnerDocument;
            }

            FusionMessageDefinitionCollectionData messageData = currentState as FusionMessageDefinitionCollectionData;

            if (messageData == null || messageData.Present == false)
            {
                switch (this.InitialState)
                {
                    case InitialState.Nill:
                        XmlAttribute attribute = doc.CreateAttribute(NilAttribute, XSISchema);
                        attribute.Value = "true";
                        messageRoot.RemoveAll();
                        messageRoot.Attributes.SetNamedItem(attribute);
                        break;

                    // Remove node from template
                    case InitialState.NotPresent:
                        messageRoot.ParentNode.RemoveChild(messageRoot);
                        break;

                    // Remove node from template
                    default:
                        messageRoot.ParentNode.RemoveChild(messageRoot);
                        break;
                }
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
