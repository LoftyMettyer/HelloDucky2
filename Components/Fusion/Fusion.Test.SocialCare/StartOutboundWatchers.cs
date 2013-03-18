

namespace Fusion.Test.SocialCare
{
    using System.IO;
    using Fusion.Core;
    using Fusion.Core.Test;
    using NServiceBus;
    using StructureMap;
    using System.Reflection;
    using log4net;
    using StructureMap.Attributes;

    public class StartOutboundWatchers : IWantToRunAtStartup
    {
            public IBus Bus { get; set; }

            private static readonly ILog Logger = LogManager.GetLogger(typeof(StartOutboundWatchers));

            [SetterProperty]
            public ITestingConfiguration TestingConfiguration
            {
                get;
                set;
            }


            public void Run()
            {
                string directory = TestingConfiguration.MessagePath;
                string communityName = TestingConfiguration.Community;

                Logger.InfoFormat("Using base directory for messages as: {0}", directory);

                DirectoryUtil.EnforceDirectory(Path.Combine(directory, "out"));
                DirectoryUtil.EnforceDirectory(Path.Combine(directory, "in"));

                OutboundWatcherDefinition[] list = new OutboundWatcherDefinition[] {
                    new OutboundWatcherDefinition { 
                        PathToWatch = "StaffChangeRequest", 
                        MessageType = typeof(Fusion.Messages.SocialCare.StaffChangeRequest),
                        MetadataExtrator = new SocialCareGenericMetadataExtractor("staffChange", "staffRef")
                    },

                    new OutboundWatcherDefinition { 
                        PathToWatch = "StaffContractChangeRequest", 
                        MessageType = typeof(Fusion.Messages.SocialCare.StaffContractChangeRequest),
                        MetadataExtrator = new SocialCareGenericMetadataExtractor("staffContractChange", "staffContractRef", "staffRef")
                    },

                    new OutboundWatcherDefinition { 
                        PathToWatch = "StaffContactChangeRequest", 
                        MessageType = typeof(Fusion.Messages.SocialCare.StaffContactChangeRequest),
                        MetadataExtrator = new SocialCareGenericMetadataExtractor("staffContactChange", "staffContactRef", "staffRef")
                    },


                    new OutboundWatcherDefinition { 
                        PathToWatch = "StaffLegalDocumentChangeRequest", 
                        MessageType = typeof(Fusion.Messages.SocialCare.StaffLegalDocumentChangeRequest),
                        MetadataExtrator = new SocialCareGenericMetadataExtractor("staffLegalDocumentChange", "staffLegalDocumentRef", "staffRef")
                    },

                    new OutboundWatcherDefinition { 
                        PathToWatch = "StaffPictureChangeRequest", 
                        MessageType = typeof(Fusion.Messages.SocialCare.StaffPictureChangeRequest),
                        MetadataExtrator = new SocialCareGenericMetadataExtractor("staffPictureChange", "staffRef")
                    },

                    new OutboundWatcherDefinition { 
                        PathToWatch = "StaffSkillChangeRequest", 
                        MessageType = typeof(Fusion.Messages.SocialCare.StaffSkillChangeRequest),
                        MetadataExtrator = new SocialCareGenericMetadataExtractor("staffSkillChange", "staffSkillRef", "staffRef")
                    },

                    new OutboundWatcherDefinition { 
                        PathToWatch = "ServiceUserChangeRequest", 
                        MessageType = typeof(Fusion.Messages.SocialCare.ServiceUserChangeRequest),
                        MetadataExtrator = new SocialCareGenericMetadataExtractor("serviceUserChange", "serviceUserRef")

                    },

                    new OutboundWatcherDefinition { 
                        PathToWatch = "ServiceUserContactChangeRequest", 
                        MessageType = typeof(Fusion.Messages.SocialCare.ServiceUserContactChangeRequest),
                        MetadataExtrator = new SocialCareGenericMetadataExtractor("serviceUserContactChange", "serviceUserContactRef", "serviceUserRef")
                    },

                    new OutboundWatcherDefinition { 
                        PathToWatch = "ServiceUserPictureChangeRequest", 
                        MessageType = typeof(Fusion.Messages.SocialCare.ServiceUserPictureChangeRequest),
                        MetadataExtrator = new SocialCareGenericMetadataExtractor("serviceUserPictureChange", "serviceUserRef")
                    },

                    new OutboundWatcherDefinition { 
                        PathToWatch = "ServiceUserHomeAddressChangeRequest", 
                        MessageType = typeof(Fusion.Messages.SocialCare.ServiceUserHomeAddressChangeRequest),
                        MetadataExtrator = new SocialCareGenericMetadataExtractor("serviceUserHomeAddressChange", "serviceUserRef")
                    },

                    new OutboundWatcherDefinition { 
                        PathToWatch = "ServiceUserCareDeilveryAddressChangeRequest", 
                        MessageType = typeof(Fusion.Messages.SocialCare.ServiceUserCareDeliveryAddressChangeRequest),
                        MetadataExtrator = new SocialCareGenericMetadataExtractor("serviceUserCareDeliveryAddressChange", "serviceUserRef")
                    },
                
                    new OutboundWatcherDefinition { 
                        PathToWatch = "ServiceUserFunderChangeRequest", 
                        MessageType = typeof(Fusion.Messages.SocialCare.ServiceUserFunderChangeRequest) ,
                        MetadataExtrator = new SocialCareGenericMetadataExtractor("serviceUserFunderChange", "serviceUserFunderRef")
                    },

                    new OutboundWatcherDefinition { 
                        PathToWatch = "ServiceUserStayChangeRequest", 
                        MessageType = typeof(Fusion.Messages.SocialCare.ServiceUserStayChangeRequest) ,
                        MetadataExtrator = new SocialCareGenericMetadataExtractor("serviceUserStayChange", "serviceUserStayRef", "serviceUserRef")
                    },

                    new OutboundWatcherDefinition { 
                        PathToWatch = "ServiceUserCareSysFundingChange", 
                        MessageType = typeof(Fusion.Messages.SocialCare.ServiceUserCareSysFundingChangeMessage),
                        MetadataExtrator = new SocialCareGenericMetadataExtractor("serviceUserCareSysFundingChange", "fundingRef")
                    },

                    new OutboundWatcherDefinition { 
                        PathToWatch = "CaresysInvoiceLineInsert", 
                        MessageType = typeof(Fusion.Messages.SocialCare.CareSysInvoiceLineInsertMessage),
                        MetadataExtrator = new SocialCareGenericMetadataExtractor("careSysInvoiceLineInsert", "CareSysInvoiceLineRef")  // case problem in message?
                    },

                    new OutboundWatcherDefinition { 
                        PathToWatch = "ServiceUserDailyRecordChange", 
                        MessageType = typeof(Fusion.Messages.SocialCare.ServiceUserDailyRecordChangeMessage) ,
                        MetadataExtrator = new SocialCareGenericMetadataExtractor("serviceUserDailyRecordChange", "serviceUserDailyRecordRef", "serviceUserRef")

                    },
                };

                string outRoot = Path.Combine(directory, "out");



                foreach(var s in list) {
                    IOutboundMessageWatcher mw = ObjectFactory.With("path").EqualTo(Path.Combine(outRoot, s.PathToWatch))
                                    .With("communityName").EqualTo(communityName)
                                    .With("metadataExtractor").EqualTo(s.MetadataExtrator)
                                    .GetInstance(typeof(OutboundMessageWatcher<>).MakeGenericType(s.MessageType)) as IOutboundMessageWatcher;

                    mw.MessageSender = ObjectFactory.GetInstance<GenericTestMessageSender>();

                    if (mw != null)
                    {
                        mw.Start();
                    }        
                }

                QuickOutboundMessageWatcher quickWatcher = new QuickOutboundMessageWatcher(outRoot);
                quickWatcher.WatcherDefinitions = list;

                quickWatcher.Start();

            }

            public void Stop()
            {

            }
    }

   
    
}
