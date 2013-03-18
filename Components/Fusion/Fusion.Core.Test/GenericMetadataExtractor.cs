using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Fusion.Messages.General;
using System.Xml.Linq;
using System.IO;

namespace Fusion.Core.Test
{
    public class GenericMetadataExtractor : IFusionXmlMetadataExtract
    {
        public GenericMetadataExtractor(string xmlNamespace, string rootNodeName, string versionNodeName, string entityRefNodeName, string primaryEntityRefNodeName)
        {
            this.xmlNamespace = xmlNamespace;
            this.rootNodeName = rootNodeName;
            this.entityRefNodeName = entityRefNodeName;
            this.primaryEntityRefNodeName = primaryEntityRefNodeName;
            this.versionNodeName = versionNodeName;
        }

        private string xmlNamespace;
        private string rootNodeName;
        private string entityRefNodeName;
        private string versionNodeName;
        private string primaryEntityRefNodeName;

        public FusionXmlMetadata GetMetadataFromXml(FusionMessage message)
        {
            XNamespace ns = this.xmlNamespace;
            XDocument fusionDocument = XDocument.Load(new StringReader(message.Xml));

            var rootNode = fusionDocument.Element(ns + this.rootNodeName);

            if (rootNode == null)
            {
                throw new ApplicationException(
                    String.Format("Cannot find message root node, expected {0}:{1}",
                        this.xmlNamespace, this.rootNodeName));
            }

            string version = (string)rootNode.Attribute(this.versionNodeName);

            if (version == null)
            {
                throw new ApplicationException(
                    String.Format("Cannot find version attribute, expected {0}/@{1}",
                        this.rootNodeName, this.versionNodeName));
            }

            Guid? entityRef = null;
            if (entityRefNodeName != null)
            {
                string entityRefString = (string)rootNode.Attribute(this.entityRefNodeName);

                if (entityRefString == null)
                {
                    throw new ApplicationException(
                        String.Format("Cannot find entity ref attribute, expected {0}/@{1}",
                            this.rootNodeName, this.entityRefNodeName));
                }

                entityRef = new Guid(entityRefString);
            }
            
            Guid? primaryEntityRef = null;

            if (primaryEntityRefNodeName != null)
            {
                string primaryEntityRefString = (string)rootNode.Attribute(this.primaryEntityRefNodeName);

                if (primaryEntityRefString == null)
                {
                    throw new ApplicationException(
                        String.Format("Cannot find primary entity ref attribute, expected {0}/@{1}",
                            this.rootNodeName, this.primaryEntityRefNodeName));
                }

                primaryEntityRef = new Guid(primaryEntityRefString);
            }
            else
            {
                primaryEntityRef = entityRef;
            }

            return new FusionXmlMetadata
            {
                EntityRef = entityRef,
                Version = Convert.ToInt32(version),
                PrimaryEntityRef = primaryEntityRef
            };
        }
    }
}
