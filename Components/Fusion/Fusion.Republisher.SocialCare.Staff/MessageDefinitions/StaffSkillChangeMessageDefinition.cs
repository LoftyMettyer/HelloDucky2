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
    public class StaffSkillChangeMessageDefinition : IMessageDefinition, IMessageValidator
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
            Stream s = Assembly.GetExecutingAssembly().GetManifestResourceStream("Fusion.Publisher.SocialCare.MessageDefinitions.templates.staffSkillChangeMessage.xml");

            return s;
        }

        public IMessageValidatorResults ValidateMessage(string xml)
        {
            FusionSchemaValidator fv = new FusionSchemaValidator();
            bool valid = fv.CheckValidity(
                xml,
                "http://advancedcomputersoftware.com/xml/fusion/socialCare",
                "res://Fusion.Messages.SocialCare/Fusion.Messages.SocialCare/xsd/",
                "staffSkillChange.xsd");

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
                                XPath = "ahc:staffSkillChange/@staffRef",
                                InitialState = InitialState.Empty,
                                Flags = XmlFlags.Mandatory
                            }
                    },

                    new FusionMessageDefinition {
                        Tag = "staffSkillRef",
                        MessageExtractor = 
                            new FusionMessageSimpleAttribute {
                                XPath = "ahc:staffSkillChange/@staffSkillRef",
                                InitialState = InitialState.Empty,
                                Flags = XmlFlags.Mandatory
                            }
                    },

                    new FusionMessageDefinition {
                        Tag = "auditUserName",
                        MessageExtractor = 
                            new FusionMessageSimpleAttribute {
                                XPath = "ahc:staffSkillChange/ahc:data/@auditUserName",
                                InitialState = InitialState.Empty,
                                Flags = XmlFlags.Mandatory
                            }
                    },

                    new FusionMessageDefinition {
                        Tag = "staffSkillData",
                        MessageExtractor =  new FusionMessageRescindableNode {
                            XPath ="ahc:staffSkillChange/ahc:data",
                            ChildNode = "ahc:staffSkill",
                            MessageData = new FusionMessageDefinition[] {                        
                                    new FusionMessageDefinition {
                                        Tag = "name",
                                        MessageExtractor = new FusionMessageSimpleNode {
                                            XPath = "ahc:name",
                                            InitialState = InitialState.Empty,
                                            Flags = XmlFlags.Mandatory
                                        }
                                    },

                                    new FusionMessageDefinition {
                                        Tag = "trainingStart",
                                        MessageExtractor = new FusionMessageSimpleNode {
                                            XPath = "ahc:trainingStart",
                                            InitialState = InitialState.NotPresent,
                                            Flags = XmlFlags.Optional
                                        }
                                    },

                                    new FusionMessageDefinition {
                                        Tag = "trainingEnd",
                                        MessageExtractor = new FusionMessageSimpleNode {
                                            XPath = "ahc:trainingEnd",
                                            InitialState = InitialState.NotPresent,
                                            Flags = XmlFlags.Optional | XmlFlags.Nillable
                                        }
                                    },

                                    new FusionMessageDefinition {
                                        Tag = "validFrom",
                                        MessageExtractor = new FusionMessageSimpleNode {
                                            XPath = "ahc:validFrom",
                                            InitialState = InitialState.NotPresent,
                                            Flags = XmlFlags.Optional
                                        }
                                    },

                                    new FusionMessageDefinition {
                                        Tag = "validTo",
                                        MessageExtractor = new FusionMessageSimpleNode {
                                            XPath = "ahc:validTo",
                                            InitialState = InitialState.NotPresent,
                                            Flags = XmlFlags.Optional | XmlFlags.Nillable
                                        }
                                    },

                                    new FusionMessageDefinition {
                                        Tag = "reference",
                                        MessageExtractor = new FusionMessageSimpleNode {
                                            XPath = "ahc:reference",
                                            InitialState = InitialState.NotPresent,
                                            Flags = XmlFlags.Optional | XmlFlags.Nillable
                                        }
                                    },

                                    new FusionMessageDefinition {
                                        Tag = "outcome",
                                        MessageExtractor = new FusionMessageSimpleNode {
                                            XPath = "ahc:outcome",
                                            InitialState = InitialState.NotPresent,
                                            Flags = XmlFlags.Nillable | XmlFlags.Optional
                                        }
                                    },

                                    new FusionMessageDefinition {
                                        Tag = "didNotAttend",
                                        MessageExtractor = new FusionMessageSimpleNode {
                                            XPath = "ahc:didNotAttend",
                                            InitialState = InitialState.NotPresent,
                                            Flags = XmlFlags.Nillable | XmlFlags.Optional
                                        }
                                    },
                                       
                                }
                            }
                        }
                    };

    }
}
