
namespace Fusion.Publisher.SocialCare.MessageDefinitions
{
    using Fusion.Republisher.Core;
    using Fusion.Republisher.Core.MessageProcessors;
    using System.IO;
    using System.Reflection;
    using System.Xml;

    public class ServiceUserStayChangeMessageDefinition : IMessageDefinition
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
            Stream s = Assembly.GetExecutingAssembly().GetManifestResourceStream("Fusion.Publisher.SocialCare.MessageDefinitions.templates.serviceUserStayChangeMessage.xml");

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
                                XPath = "ahc:serviceUserStayChange/@serviceUserRef",
                                InitialState = InitialState.Empty,
                                Flags = XmlFlags.Mandatory
                            }
                    },

                   new FusionMessageDefinition {
                        Tag = "serviceUserStayRef",
                        MessageExtractor = 
                            new FusionMessageSimpleAttribute {
                                XPath = "ahc:serviceUserStayChange/@serviceUserStayRef",
                                InitialState = InitialState.Empty,
                                Flags = XmlFlags.Mandatory
                            }
                    },


                    new FusionMessageDefinition {
                        Tag = "auditUserName",
                        MessageExtractor = 
                            new FusionMessageSimpleAttribute {
                                XPath = "ahc:serviceUserStayChange/ahc:data/@auditUserName",
                                InitialState = InitialState.Empty,
                                Flags = XmlFlags.Mandatory
                            }
                    },

                    new FusionMessageDefinition {
                        Tag = "serviceUserStayData",
                        MessageExtractor =  new FusionMessageRescindableNode {
                            XPath ="ahc:serviceUserStayChange/ahc:data",
                            ChildNode = "ahc:serviceUserStay",
                            MessageData = new FusionMessageDefinition[] {                        
                                    new FusionMessageDefinition {
                                        Tag = "primarySite",
                                        MessageExtractor = new FusionMessageSimpleNode {
                                            XPath = "ahc:primarySite",
                                            InitialState = InitialState.Nill,
                                            Flags = XmlFlags.Mandatory | XmlFlags.Nillable
                                        }
                                    },

                                    new FusionMessageDefinition {
                                        Tag = "room",
                                        MessageExtractor = new FusionMessageSimpleNode {
                                            XPath = "ahc:room",
                                            InitialState = InitialState.NotPresent,
                                            Flags = XmlFlags.Optional | XmlFlags.Nillable
                                        }
                                    },

                                    new FusionMessageDefinition {
                                        Tag = "admissionDate",
                                        MessageExtractor = new FusionMessageSimpleNode {
                                            XPath = "ahc:admissionDate",
                                            InitialState = InitialState.Empty,
                                            Flags = XmlFlags.Mandatory
                                        }
                                    },

                                    new FusionMessageDefinition {
                                        Tag = "dischargeDate",
                                        MessageExtractor = new FusionMessageSimpleNode {
                                            XPath = "ahc:dischargeDate",
                                            InitialState = InitialState.NotPresent,
                                            Flags = XmlFlags.Optional | XmlFlags.Nillable
                                        }
                                    },

                                    new FusionMessageDefinition {
                                        Tag = "serviceType",
                                        MessageExtractor = new FusionMessageSimpleNode {
                                            XPath = "ahc:serviceType",
                                            InitialState = InitialState.Value,
                                            Flags = XmlFlags.Mandatory
                                        }
                                    },


                                    new FusionMessageDefinition {
                                        Tag = "fundingType",
                                        MessageExtractor = new FusionMessageSimpleNode {
                                            XPath = "ahc:fundingType",
                                            InitialState = InitialState.Value,
                                            Flags = XmlFlags.Mandatory
                                        }
                                    },


                                    new FusionMessageDefinition {
                                        Tag = "locationWithinFacility",
                                        MessageExtractor = 
                                            new FusionMessageSimpleNode {
                                                XPath = "ahc:locationWithinFacility",
                                                InitialState = InitialState.NotPresent,
                                                Flags = XmlFlags.Optional | XmlFlags.Nillable
                                            }
                                    },

                                    new FusionMessageDefinition {
                                        Tag = "companyName",
                                        MessageExtractor = 
                                            new FusionMessageSimpleNode {
                                                XPath = "ahc:companyName",
                                                InitialState = InitialState.NotPresent,
                                                Flags = XmlFlags.Optional | XmlFlags.Nillable
                                            }
                                    },
                                    
                                              
                                }
                            }
                        }
                    };

    }
}
