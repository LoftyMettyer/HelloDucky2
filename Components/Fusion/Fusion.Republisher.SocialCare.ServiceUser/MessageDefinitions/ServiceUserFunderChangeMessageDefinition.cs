using Fusion.Republisher.Core;
using Fusion.Republisher.Core.MessageProcessors;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Xml;

namespace Fusion.Publisher.SocialCare.MessageDefinitions
{
    public class ServiceUserFunderChangeMessageDefinition : IMessageDefinition
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
            Stream s = Assembly.GetExecutingAssembly().GetManifestResourceStream("Fusion.Publisher.SocialCare.MessageDefinitions.templates.serviceUserFunderChangeMessage.xml");

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
                        Tag = "serviceUserFunderRef",
                        MessageExtractor = 
                            new FusionMessageSimpleAttribute {
                                XPath = "ahc:serviceUserFunderChange/@serviceUserFunderRef",
                                InitialState = InitialState.Empty,
                                Flags = XmlFlags.Mandatory
                            }
                    },

                    new FusionMessageDefinition {
                        Tag = "auditUserName",
                        MessageExtractor = 
                            new FusionMessageSimpleAttribute {
                                XPath = "ahc:serviceUserFunderChange/ahc:data/@auditUserName",
                                InitialState = InitialState.Empty,
                                Flags = XmlFlags.Mandatory
                            }
                    },

                    new FusionMessageDefinition {
                        Tag = "serviceUserFunderData",
                        MessageExtractor =  new FusionMessageRescindableNode {
                            XPath ="ahc:serviceUserFunderChange/ahc:data",
                            ChildNode = "ahc:serviceUserFunder",
                            MessageData = new FusionMessageDefinition[] {                        
                                    new FusionMessageDefinition {
                                        Tag = "accountCode",
                                        MessageExtractor = new FusionMessageSimpleNode {
                                            XPath = "ahc:accountCode",
                                            InitialState = InitialState.Value,
                                            Flags = XmlFlags.Mandatory
                                        }
                                    },

                                    new FusionMessageDefinition {
                                        Tag = "accountLookup",
                                        MessageExtractor = new FusionMessageSimpleNode {
                                            XPath = "ahc:accountLookup",
                                            InitialState = InitialState.Nill,
                                            Flags = XmlFlags.Optional | XmlFlags.Nillable
                                        }
                                    },

                                    new FusionMessageDefinition {
                                        Tag = "costCenter",
                                        MessageExtractor = new FusionMessageSimpleNode {
                                            XPath = "ahc:costCenter",
                                            InitialState = InitialState.Value,
                                            Flags = XmlFlags.Mandatory
                                        }
                                    },

                                    new FusionMessageDefinition {
                                        Tag = "site",
                                        MessageExtractor = new FusionMessageSimpleNode {
                                            XPath = "ahc:site",
                                            InitialState = InitialState.Value,
                                            Flags = XmlFlags.Mandatory
                                        }
                                    },

                                    new FusionMessageDefinition {
                                        Tag = "name",
                                        MessageExtractor = new FusionMessageSimpleNode {
                                            XPath = "ahc:name",
                                            InitialState = InitialState.Value,
                                            Flags = XmlFlags.Mandatory
                                        }
                                    },

                                    new FusionMessageDefinition {
                                        Tag = "addressLine1",
                                        MessageExtractor = new FusionMessageSimpleNode {
                                            XPath = "ahc:addressLine1",
                                            InitialState = InitialState.Empty,
                                            Flags = XmlFlags.Mandatory
                                        }
                                    },

                                    new FusionMessageDefinition {
                                        Tag = "addressLine2",
                                        MessageExtractor = new FusionMessageSimpleNode {
                                            XPath = "ahc:addressLine2",
                                            InitialState = InitialState.Empty,
                                            Flags = XmlFlags.Mandatory
                                        }
                                    },

                                    new FusionMessageDefinition {
                                        Tag = "addressLine3",
                                        MessageExtractor = new FusionMessageSimpleNode {
                                            XPath = "ahc:addressLine3",
                                            InitialState = InitialState.Empty,
                                            Flags = XmlFlags.Mandatory
                                        }
                                    },

                                    new FusionMessageDefinition {
                                        Tag = "addressLine4",
                                        MessageExtractor = new FusionMessageSimpleNode {
                                            XPath = "ahc:addressLine4",
                                            InitialState = InitialState.Empty,
                                            Flags = XmlFlags.Mandatory
                                        }
                                    },

                                    new FusionMessageDefinition {
                                        Tag = "addressLine5",
                                        MessageExtractor = new FusionMessageSimpleNode {
                                            XPath = "ahc:addressLine5",
                                            InitialState = InitialState.Empty,
                                            Flags = XmlFlags.Mandatory
                                        }
                                    },

                                    new FusionMessageDefinition {
                                        Tag = "postCode",
                                        MessageExtractor = new FusionMessageSimpleNode {
                                            XPath = "ahc:postCode",
                                            InitialState = InitialState.Empty,
                                            Flags = XmlFlags.Mandatory
                                        }
                                    },

                                    new FusionMessageDefinition {
                                        Tag = "phoneNumber",
                                        MessageExtractor = new FusionMessageSimpleNode {
                                            XPath = "ahc:workMobile",
                                            InitialState = InitialState.NotPresent,
                                            Flags = XmlFlags.Nillable | XmlFlags.Optional
                                        }
                                    },

                                    new FusionMessageDefinition {
                                        Tag = "email",
                                        MessageExtractor = new FusionMessageSimpleNode {
                                            XPath = "ahc:email",
                                            InitialState = InitialState.NotPresent,
                                            Flags = XmlFlags.Nillable | XmlFlags.Optional
                                        }
                                    },

                                    new FusionMessageDefinition {
                                        Tag = "paymentMethod",
                                        MessageExtractor = new FusionMessageSimpleNode {
                                            XPath = "ahc:paymentMethod",
                                            InitialState = InitialState.Value,
                                            Flags = XmlFlags.Mandatory
                                        }
                                    },

                                    new FusionMessageDefinition {
                                        Tag = "paymentDays",
                                        MessageExtractor = new FusionMessageSimpleNode {
                                            XPath = "ahc:paymentDays",
                                            InitialState = InitialState.Value,
                                            Flags = XmlFlags.Mandatory
                                        }
                                    },

                                    new FusionMessageDefinition {
                                        Tag = "sendDocumentsBy",
                                        MessageExtractor = new FusionMessageSimpleNode {
                                            XPath = "ahc:sendDocumentsBy",
                                            InitialState = InitialState.Value,
                                            Flags = XmlFlags.Mandatory
                                        }
                                    },

                                    new FusionMessageDefinition {
                                        Tag = "accountOnHold",
                                        MessageExtractor = new FusionMessageSimpleNode {
                                            XPath = "ahc:accountOnHold",
                                            InitialState = InitialState.Value,
                                            Flags = XmlFlags.Mandatory
                                        }
                                    },

                                    new FusionMessageDefinition {
                                        Tag = "analysisCode",
                                        MessageExtractor = new FusionMessageSimpleNode {
                                            XPath = "ahc:v",
                                            InitialState = InitialState.NotPresent,
                                            Flags = XmlFlags.Optional | XmlFlags.Nillable
                                        }
                                    },

                                    new FusionMessageDefinition {
                                        Tag = "accountGrouping",
                                        MessageExtractor = new FusionMessageSimpleNode {
                                            XPath = "ahc:paymentMethod",
                                            InitialState = InitialState.NotPresent,
                                            Flags = XmlFlags.Optional | XmlFlags.Nillable
                                        }
                                    },

                                    new FusionMessageDefinition {
                                        Tag = "defaultNominalCode",
                                        MessageExtractor = new FusionMessageSimpleNode {
                                            XPath = "ahc:defaultNominalCode",
                                            InitialState = InitialState.NotPresent,
                                            Flags = XmlFlags.Nillable | XmlFlags.Optional
                                        }
                                    },

                                    new FusionMessageDefinition {
                                        Tag = "companyName",
                                        MessageExtractor = new FusionMessageSimpleNode {
                                            XPath = "ahc:companyName",
                                            InitialState = InitialState.NotPresent,
                                            Flags = XmlFlags.Optional
                                        }
                                    },

                                    new FusionMessageDefinition {
                                        Tag = "multipleSUFunder",
                                        MessageExtractor = new FusionMessageSimpleNode {
                                            XPath = "ahc:multipleSUFunder",
                                            InitialState = InitialState.Value,
                                            Flags = XmlFlags.Mandatory
                                        }
                                    },
                                              

                                }
                            }
                        }
                    };

    }
}
