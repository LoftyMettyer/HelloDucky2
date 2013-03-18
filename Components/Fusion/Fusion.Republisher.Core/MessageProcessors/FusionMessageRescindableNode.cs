using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace Fusion.Republisher.Core.MessageProcessors
{
    public class FusionMessageRescindableNode : IExtractFusionMessageData
    {
        public string XPath;
        public string ChildNode;
        public FusionMessageDefinition[] MessageData;

        const string RecordStatusAttributeName = "recordStatus";

        public object ReadFromMessage(System.Xml.XmlNode message, System.Xml.XmlNamespaceManager ns, object currentState)
        {
            FusionMessageRescindableNodeData rescindableNodeData = currentState as FusionMessageRescindableNodeData;
            if (rescindableNodeData == null)
            {
                rescindableNodeData = new FusionMessageRescindableNodeData
                {
                    Data = new Dictionary<string, object>()
                };
            }
            XmlNode node = message.SelectSingleNode(this.XPath, ns);

            XmlAttribute recordStatusAttribute = node.Attributes[RecordStatusAttributeName];
            if (recordStatusAttribute != null)
            {
                string recordStatus = recordStatusAttribute.InnerText;

                switch (recordStatus)
                {
                    case "Active":
                    case "Inactive":
                        rescindableNodeData.Status = recordStatus;

                        FusionMessageDefinitionCollection collection = new FusionMessageDefinitionCollection
                        {
                            MessageData = this.MessageData,
                            XPath = ChildNode
                        };

                        var data = collection.ReadFromMessage(node, ns, new FusionMessageDefinitionCollectionData
                        {
                            Data = rescindableNodeData.Data ?? new Dictionary<string, object>()
                        });

                        rescindableNodeData.Data = ((FusionMessageDefinitionCollectionData)data).Data;

                        return rescindableNodeData;
                        
                    case "RecordCreatedInError":
                        return new FusionMessageRescindableNodeData
                        {
                            Status = recordStatus,
                            Data = null
                        };
                }
            }

            return null;
        }
            

        public void Update(System.Xml.XmlNode message, System.Xml.XmlNamespaceManager ns, object messageData)
        {
            FusionMessageRescindableNodeData rescindableNodeData = messageData as FusionMessageRescindableNodeData;

            if (messageData == null)
            {
                // can't do anything!
                return;
            }

            XmlNode node = message.SelectSingleNode(this.XPath, ns);

            if (node == null)
            {
                // can't find null
            }

            XmlDocument doc = message as XmlDocument;
            if (doc == null)
            {
                doc = message.OwnerDocument;
            }

            string recordStatus = rescindableNodeData.Status;

            // Set attribute

            XmlAttribute attribute = doc.CreateAttribute(RecordStatusAttributeName);
            attribute.Value = recordStatus;
            node.Attributes.SetNamedItem(attribute);

            switch (recordStatus)
            {
                case "Active":
                case "Inactive":

                    FusionMessageDefinitionCollection collection = new FusionMessageDefinitionCollection
                    {
                        MessageData = this.MessageData,
                        XPath = ChildNode
                    };

                    collection.Update(node, ns, new FusionMessageDefinitionCollectionData
                                    {
                                            Present = true,
                                            Data = rescindableNodeData.Data
                                    });

                   break;

                case "RecordCreatedInError":
                    XmlNode childNode = node.SelectSingleNode(this.ChildNode, ns);
                    if (childNode != null)
                    {
                        node.RemoveChild(childNode);
                    }
                    break;
            }

        }
    }

    public class FusionMessageRescindableNodeData
    {
        public string Status;
        public Dictionary<string, object> Data;
    }
    
}
