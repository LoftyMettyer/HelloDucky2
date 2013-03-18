
namespace Fusion.Publisher.SocialCare.MessageDefinitions
{
    using Fusion.Republisher.Core;
    using Fusion.Republisher.Core.MessageProcessors;
    using System.IO;
    using System.Reflection;
    using System.Xml;

    public class ServiceUserChangeMessageDefinition : IMessageDefinition
    {
        public FusionMessageDefinition[] MessageDefinition
        {
            get
            {
                return descriptionMap;
            }
        }

        public Stream GetBlankXmlTemplate()
        {
            Stream s = Assembly.GetExecutingAssembly().GetManifestResourceStream("Fusion.Publisher.SocialCare.MessageDefinitions.templates.serviceUserChangeMessage.xml");

            return s;
        }

        private XmlNamespaceManager ns;

        public XmlNamespaceManager NamespaceManager
        {
            get
            {
                if (ns != null)
                {
                    return ns;
                }

                ns = new XmlNamespaceManager(new NameTable());
                ns.AddNamespace("ahc", "http://advancedcomputersoftware.com/xml/fusion/socialCare");

                return ns;
            }
        }

        private readonly FusionMessageDefinition[] descriptionMap = new FusionMessageDefinition[] {
                   new FusionMessageDefinition {
                        Tag = "serviceUserRef",
                        MessageExtractor = 
                            new FusionMessageSimpleAttribute {
                                XPath = "ahc:serviceUserChange/@serviceUserRef",
                                InitialState = InitialState.Empty,
                                Flags = XmlFlags.Mandatory
                            }
                    },

                    new FusionMessageDefinition {
                        Tag = "auditUserName",
                        MessageExtractor = 
                            new FusionMessageSimpleAttribute {
                                XPath = "ahc:serviceUserChange/ahc:data/@auditUserName",
                                InitialState = InitialState.Empty,
                                Flags = XmlFlags.Mandatory
                            }
                    },

                    new FusionMessageDefinition {
                        Tag = "serviceUserData",
                        MessageExtractor =  new FusionMessageRescindableNode {
                            XPath ="ahc:serviceUserChange/ahc:data",
                            ChildNode = "ahc:serviceUser",
                            MessageData = new FusionMessageDefinition[] {                        
                                    new FusionMessageDefinition {
                                        Tag = "title",
                                        MessageExtractor = new FusionMessageSimpleNode {
                                            XPath = "ahc:title",
                                            InitialState = InitialState.Empty,
                                            Flags = XmlFlags.Mandatory
                                        }
                                    },

                                    new FusionMessageDefinition {
                                        Tag = "forenames",
                                        MessageExtractor = new FusionMessageSimpleNode {
                                            XPath = "ahc:forenames",
                                            InitialState = InitialState.Empty,
                                            Flags = XmlFlags.Mandatory
                                        }
                                    },

                                    new FusionMessageDefinition {
                                        Tag = "surname",
                                        MessageExtractor = new FusionMessageSimpleNode {
                                            XPath = "ahc:surname",
                                            InitialState = InitialState.Empty,
                                            Flags = XmlFlags.Mandatory
                                        }
                                    },

                                    new FusionMessageDefinition {
                                        Tag = "DOB",
                                        MessageExtractor = new FusionMessageSimpleNode {
                                            XPath = "ahc:DOB",
                                            InitialState = InitialState.Value,
                                            Flags = XmlFlags.Mandatory
                                        }
                                    },

                                    new FusionMessageDefinition {
                                        Tag = "gender",
                                        MessageExtractor = new FusionMessageSimpleNode {
                                            XPath = "ahc:gender",
                                            InitialState = InitialState.Value,
                                        }
                                    },


                                    new FusionMessageDefinition {
                                        Tag = "serviceUserNumber",
                                        MessageExtractor = new FusionMessageSimpleNode {
                                            XPath = "ahc:serviceUserNumber",
                                            InitialState = InitialState.Empty,
                                            Flags = XmlFlags.Optional
                                        }
                                    },


                                    new FusionMessageDefinition {
                                        Tag = "companyName",
                                        MessageExtractor = 
                                            new FusionMessageSimpleNode {
                                                XPath = "ahc:companyName",
                                                InitialState = InitialState.NotPresent,
                                                Flags = XmlFlags.Optional
                                            }
                                    },
                                                                                  
                                }
                            }
                        }
                    };

    }
}
