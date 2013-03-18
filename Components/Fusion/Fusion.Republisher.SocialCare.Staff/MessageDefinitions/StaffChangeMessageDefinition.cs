
namespace Fusion.Publisher.SocialCare.MessageDefinitions
{
    using Fusion.Republisher.Core;
    using Fusion.Republisher.Core.MessageProcessors;
    using System.IO;
    using System.Reflection;
    using System.Xml;

    public class StaffChangeMessageDefinition : IMessageDefinition, IMessageValidator
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
            Stream s = Assembly.GetExecutingAssembly().GetManifestResourceStream("Fusion.Publisher.SocialCare.MessageDefinitions.templates.staffChangeMessage.xml");

            return s;
        }

        public IMessageValidatorResults ValidateMessage(string xml)
        {
            FusionSchemaValidator fv = new FusionSchemaValidator();
            bool valid = fv.CheckValidity(
                xml, 
                "http://advancedcomputersoftware.com/xml/fusion/socialCare",
                "res://Fusion.Messages.SocialCare/Fusion.Messages.SocialCare/xsd/",
                "staffChange.xsd"); 
              
            string validationMessage = fv.ValidationMessage;

            return new MessageValidationResults
            {
                IsValid = valid,
                ValidationMessage = validationMessage
            };
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
                        Tag = "staffRef",
                        MessageExtractor = 
                            new FusionMessageSimpleAttribute {
                                XPath = "ahc:staffChange/@staffRef",
                                InitialState = InitialState.Empty,
                                Flags = XmlFlags.Mandatory
                            }
                    },

                    new FusionMessageDefinition {
                        Tag = "auditUserName",
                        MessageExtractor = 
                            new FusionMessageSimpleAttribute {
                                XPath = "ahc:staffChange/ahc:data/@auditUserName",
                                InitialState = InitialState.Empty,
                                Flags = XmlFlags.Mandatory
                            }
                    },

                    new FusionMessageDefinition {
                        Tag = "staffData",
                        MessageExtractor =  new FusionMessageRescindableNode {
                            XPath ="ahc:staffChange/ahc:data",
                            ChildNode = "ahc:staff",
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
                                            InitialState = InitialState.Value,
                                            Flags = XmlFlags.Mandatory
                                        }
                                    },

                                    new FusionMessageDefinition {
                                        Tag = "surname",
                                        MessageExtractor = new FusionMessageSimpleNode {
                                            XPath = "ahc:surname",
                                            InitialState = InitialState.Value,
                                            Flags = XmlFlags.Mandatory
                                        }
                                    },

                                    new FusionMessageDefinition {
                                        Tag = "preferredName",
                                        MessageExtractor = new FusionMessageSimpleNode {
                                            XPath = "ahc:preferredName",
                                            InitialState = InitialState.NotPresent,
                                            Flags = XmlFlags.Nillable | XmlFlags.Optional
                                        }
                                    },

                                    new FusionMessageDefinition {
                                        Tag = "payrollNumber",
                                        MessageExtractor = new FusionMessageSimpleNode {
                                            XPath = "ahc:payrollNumber",
                                            InitialState = InitialState.NotPresent,
                                            Flags = XmlFlags.Optional
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
                                        Tag = "employeeType",
                                        MessageExtractor = new FusionMessageSimpleNode {
                                            XPath = "ahc:employeeType",
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
                                        Tag = "personalEmail",
                                        MessageExtractor = new FusionMessageSimpleNode {
                                            XPath = "ahc:personalEmail",
                                            InitialState = InitialState.NotPresent,
                                            Flags = XmlFlags.Nillable | XmlFlags.Optional
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
                                        Tag = "startDate",
                                        MessageExtractor = new FusionMessageSimpleNode {
                                            XPath = "ahc:startDate",
                                            InitialState = InitialState.Value,
                                        }
                                    },

                                    new FusionMessageDefinition {
                                        Tag = "leavingDate",
                                        MessageExtractor = 
                                            new FusionMessageSimpleNode {
                                                XPath = "ahc:leavingDate",
                                                InitialState = InitialState.NotPresent,
                                                Flags = XmlFlags.Nillable | XmlFlags.Optional
                                            }
                                    },

                                    new FusionMessageDefinition {
                                        Tag = "leavingReason",
                                        MessageExtractor = 
                                            new FusionMessageSimpleNode {
                                                XPath = "ahc:leavingReason",
                                                InitialState = InitialState.NotPresent,
                                                Flags = XmlFlags.Nillable | XmlFlags.Optional
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

                                    new FusionMessageDefinition {
                                        Tag = "jobTitle",
                                        MessageExtractor = 
                                            new FusionMessageSimpleNode {
                                                XPath = "ahc:jobTitle",
                                                InitialState = InitialState.NotPresent,
                                                Flags = XmlFlags.Optional
                                            }
                                    },

                                    new FusionMessageDefinition {
                                        Tag = "managerRef",
                                        MessageExtractor = 
                                            new FusionMessageSimpleNode {
                                                XPath = "ahc:managerRef",
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

                                            
                                        
                                    },

                                    new FusionMessageDefinition {
                                        Tag = "nationalInsuranceNumber",
                                        MessageExtractor = 
                                            new FusionMessageSimpleNode {
                                                XPath = "ahc:nationalInsuranceNumber",
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
