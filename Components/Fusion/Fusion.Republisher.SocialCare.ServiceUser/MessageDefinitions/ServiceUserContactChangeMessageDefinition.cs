﻿using Fusion.Republisher.Core;
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
    public class ServiceUserContactChangeMessageDefinition : IMessageDefinition
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
            Stream s = Assembly.GetExecutingAssembly().GetManifestResourceStream("Fusion.Publisher.SocialCare.MessageDefinitions.templates.serviceUserContactChangeMessage.xml");

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
                        Tag = "serviceUserContactRef",
                        MessageExtractor = 
                            new FusionMessageSimpleAttribute {
                                XPath = "ahc:serviceUserContactChange/@serviceUserContactRef",
                                InitialState = InitialState.Empty,
                                Flags = XmlFlags.Mandatory
                            }
                    },

                    new FusionMessageDefinition {
                        Tag = "serviceUserRef",
                        MessageExtractor = 
                            new FusionMessageSimpleAttribute {
                                XPath = "ahc:serviceUserContactChange/@serviceUserRef",
                                InitialState = InitialState.Empty,
                                Flags = XmlFlags.Mandatory
                            }
                    },


                    new FusionMessageDefinition {
                        Tag = "auditUserName",
                        MessageExtractor = 
                            new FusionMessageSimpleAttribute {
                                XPath = "ahc:serviceUserContactChange/ahc:data/@auditUserName",
                                InitialState = InitialState.Empty,
                                Flags = XmlFlags.Mandatory
                            }
                    },

                    new FusionMessageDefinition {
                        Tag = "serviceUserContactData",
                        MessageExtractor =  new FusionMessageRescindableNode {
                            XPath ="ahc:serviceUserContactChange/ahc:data",
                            ChildNode = "ahc:serviceUserContact",
                            MessageData = new FusionMessageDefinition[] {                        
                                    new FusionMessageDefinition {
                                        Tag = "title",
                                        MessageExtractor = new FusionMessageSimpleNode {
                                            XPath = "ahc:title",
                                            InitialState = InitialState.Nill,
                                            Flags = XmlFlags.Mandatory | XmlFlags.Nillable
                                        }
                                    },

                                    new FusionMessageDefinition {
                                        Tag = "forenames",
                                        MessageExtractor = new FusionMessageSimpleNode {
                                            XPath = "ahc:forenames",
                                            InitialState = InitialState.Nill,
                                            Flags = XmlFlags.Mandatory | XmlFlags.Nillable
                                        }
                                    },

                                    new FusionMessageDefinition {
                                        Tag = "surname",
                                        MessageExtractor = new FusionMessageSimpleNode {
                                            XPath = "ahc:surname",
                                            InitialState = InitialState.Nill,
                                            Flags = XmlFlags.Mandatory | XmlFlags.Nillable
                                        }
                                    },

                                    new FusionMessageDefinition {
                                        Tag = "description",
                                        MessageExtractor = new FusionMessageSimpleNode {
                                            XPath = "ahc:description",
                                            InitialState = InitialState.NotPresent,
                                            Flags = XmlFlags.Nillable | XmlFlags.Optional
                                        }
                                    },

                                    new FusionMessageDefinition {
                                        Tag = "relationshipType",
                                        MessageExtractor = new FusionMessageSimpleNode {
                                            XPath = "ahc:relationshipType",
                                            InitialState = InitialState.Value,
                                            Flags = XmlFlags.Mandatory
                                        }
                                    },

                                    new FusionMessageDefinition {
                                        Tag = "workMobile",
                                        MessageExtractor = new FusionMessageSimpleNode {
                                            XPath = "ahc:workMobile",
                                            InitialState = InitialState.NotPresent,
                                            Flags = XmlFlags.Nillable | XmlFlags.Optional
                                        }
                                    },

                                    new FusionMessageDefinition {
                                        Tag = "personalMobile",
                                        MessageExtractor = new FusionMessageSimpleNode {
                                            XPath = "ahc:personalMobile",
                                            InitialState = InitialState.NotPresent,
                                            Flags = XmlFlags.Nillable | XmlFlags.Optional
                                        }
                                    },

                                    new FusionMessageDefinition {
                                        Tag = "workPhoneNumber",
                                        MessageExtractor = new FusionMessageSimpleNode {
                                            XPath = "ahc:workPhoneNumber",
                                            InitialState = InitialState.NotPresent,
                                            Flags = XmlFlags.Nillable | XmlFlags.Optional
                                        }
                                    },

                                    new FusionMessageDefinition {
                                        Tag = "homePhoneNumber",
                                        MessageExtractor = new FusionMessageSimpleNode {
                                            XPath = "ahc:homePhoneNumber",
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
                                        Tag = "notes",
                                        MessageExtractor = new FusionMessageSimpleNode {
                                            XPath = "ahc:notes",
                                            InitialState = InitialState.NotPresent,
                                            Flags = XmlFlags.Nillable | XmlFlags.Optional
                                        }
                                    },

                                    new FusionMessageDefinition {
                                        Tag = "homeAddress",
                                        MessageExtractor = new FusionMessageDefinitionCollection {
                                            XPath = "ahc:homeAddress",
                                            MessageData = new FusionMessageDefinition[] {
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

                                            }
                                        }
                                                                                  
                                    }

                                }
                            }
                        }
                    };

    }
}
