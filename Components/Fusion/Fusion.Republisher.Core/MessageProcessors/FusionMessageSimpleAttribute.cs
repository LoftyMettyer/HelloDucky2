using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace Fusion.Republisher.Core.MessageProcessors
{
    public class FusionMessageSimpleAttribute : IExtractFusionMessageData
    {
        const string XSISchema = "http://www.w3.org/2001/XMLSchema-instance";
        const string NilAttribute = "nil";

        public string XPath;
        public XmlFlags Flags;
        public InitialState InitialState;

        
        public object ReadFromMessage(XmlNode message, XmlNamespaceManager ns, object currentState)
        {
            XmlDocument doc = message as XmlDocument;
            if (doc == null)
            {
                doc = message.OwnerDocument;
            }

            XmlNode node = message.SelectSingleNode(this.XPath, ns);

            if (node == null)
            {
                // Node is not part of incoming message, make no change to state

                return currentState ?? new FusionMessageSimpleAttributeData
                {
                    Present = false,
                    Value = null
                };
            }
            else
            {
                // Node should have data              

                return new FusionMessageSimpleAttributeData
                {
                    Present = true,
                    Value = node.InnerText
                };
            }
        }

        public void Update(XmlNode message, XmlNamespaceManager ns, object currentValue)
        {
            XmlDocument doc = message as XmlDocument;
            if (doc == null)
            {
                doc = message.OwnerDocument;
            }

            XmlNode node = message.SelectSingleNode(this.XPath, ns);

            if (node == null)
            {
                // Cannot find node - ignore
                return;
            }

            FusionMessageSimpleAttributeData nodeData = currentValue as FusionMessageSimpleAttributeData;

            if (nodeData == null || nodeData.Present == false)
            {
                // No value to store

                switch (this.InitialState)
                {
                    case InitialState.Empty:
                        // Leave as per template
                        break;
                    case InitialState.Nill:
                        throw new NotSupportedException();
                    case InitialState.NotPresent:
                        node.ParentNode.RemoveChild(node);
                        break;
                }
            }
            else
            {
                node.InnerText = nodeData.Value;
            }
        }

    }

    public class FusionMessageSimpleAttributeData
    {
        public bool Present;
        public string Value;
    }

}
