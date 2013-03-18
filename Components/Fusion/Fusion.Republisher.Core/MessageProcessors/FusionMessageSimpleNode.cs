using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace Fusion.Republisher.Core.MessageProcessors
{
    public class FusionMessageSimpleNode : IExtractFusionMessageData
    {
        const string XSISchema = "http://www.w3.org/2001/XMLSchema-instance";
        const string NilAttribute = "nil";

        public string XPath;
        public XmlFlags Flags;
        public InitialState InitialState;

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
                // Node is not part of template, ignore
                return;
            }

            FusionMessageSimpleNodeData nodeData = currentValue as FusionMessageSimpleNodeData;

            if (nodeData == null || nodeData.Present == false)
            {
                // No value to store

                switch (this.InitialState)
                {
                    case InitialState.Empty:
                        // Leave as per template
                        break;
                    case InitialState.Nill:
                        XmlAttribute attribute = doc.CreateAttribute(NilAttribute, XSISchema);
                        attribute.Value = "true";
                        node.Attributes.SetNamedItem(attribute);
                        break;
                    case InitialState.NotPresent:
                        node.ParentNode.RemoveChild(node);
                        break;
                }
            }
            else
            {
                if (nodeData.Value == null)
                {
                    XmlAttribute attribute = doc.CreateAttribute(NilAttribute, XSISchema);
                    attribute.Value = "true";
                    node.Attributes.SetNamedItem(attribute);
                }
                else
                {
                    node.InnerText = nodeData.Value;
                }
            }
        }

        public object ReadFromMessage(XmlNode message, XmlNamespaceManager ns, object currentState)
        {
            XmlDocument doc = message as XmlDocument;
            if (doc == null) {
                doc = message.OwnerDocument;
            }
            
            XmlNode node = message.SelectSingleNode(this.XPath, ns);

            if (node == null)
            {
                // Node is not part of incoming message, make no change to state

                return currentState ?? new FusionMessageSimpleNodeData
                {
                    Present = false,
                    Value = null
                };
            }
            else
            {
                string nodeValue = null;

                if (node.Attributes[NilAttribute, XSISchema] != null)
                {
                    // Node is 'nill' - ie empty
                    nodeValue = null;
                }
                else
                {
                    // Node should have data
                    nodeValue = node.InnerText;
                }

                return new FusionMessageSimpleNodeData
                {
                    Present = true,
                    Value = nodeValue
                };
            }
        }
    }

    public class FusionMessageSimpleNodeData
    {
        public bool Present;
        public string Value;
    }

    public enum InitialState  {
        NotPresent,
        Empty,
        Nill,
        Value
    }

    [Flags]
    public enum XmlFlags : int
    {
        Mandatory = 0,
        Nillable = 0x01,
        Optional = 0x02
    }
}
